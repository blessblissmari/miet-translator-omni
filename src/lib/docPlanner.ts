/**
 * Document planning — translates extracted pages into DocPlan for DOCX building.
 */
import { chat } from "./mimo";
import { normalizeMath } from "./mathNormalize";
import { downsampleDataUrl } from "./imageOps";
import { mapWithConcurrency } from "./concurrency";
import {
  harvestPairs,
  glossaryPrompt,
  mergeGlossary,
  dspGlossaryPrompt,
  applyGlossaryPost,
  type Glossary,
} from "./glossary";
import { polishRu } from "./ruPolish";
import {
  stripCodeFences,
  parseMarkdownToBlocks,
  normalizeMathDelims,
  wrapOrphanLatex,
  sanitizeHtml,
  TARGET_LANG,
  type PlannerOpts,
} from "./plannerShared";
import type { DocPlan, DocBlock, ExtractedDoc } from "./types";
import { detectFigures, cropFigures } from "./figureDetect";
import { buildDocumentBlueprint, economyVisionModel, pageBlueprintHint, type DocumentBlueprint } from "./layoutBlueprint";

/* ─── Prompts ──────────────────────────────────── */

const VISION_OCR_PROMPT = (lang: string) =>
  `You are a senior technical translator and OCR expert for academic notes, including HANDWRITTEN material.

Task: look at the attached image of a page (it may be handwritten lecture notes, a scanned printed page, or a photo of someone's notebook). Carefully read the contents — including handwriting, formulas, sketches, and any printed text. Then translate everything into ${lang} ("русский") in academic МИЭТ-style.

CRITICAL rules:
- Output ONLY translated Markdown. No commentary. No code fences.
- **NEVER output HTML.** No <sub>, <sup>, <i>, <b>, <br>, <table>, <math>, <span>, <p>, <div> tags. Use ONLY plain Markdown and LaTeX math.
- Read the page exhaustively. Do NOT skip handwritten margin notes, sub-questions, or formulas.
- For mathematical content, use LaTeX in $...$ (inline) and $$...$$ (display). Reproduce subscripts, superscripts, fractions, integrals, sums faithfully. Multi-line environments (cases, align, matrix) MUST be wrapped in $$ ... $$. Never use \\( \\) or \\[ \\].
- **EVERY math expression must be wrapped in $...$ — even single variables like $x_1$, $V_T$, $y[n]$. Whole equations must be one math span, not fragments:
    WRONG: \`y[n] = \\mathcal{H}\\{x[n]\\}\`
    RIGHT: \`$y[n] = \\mathcal{H}\\{x[n]\\}$\`
- Use Markdown structure: # for top heading, ##/### for sections, "- item" for bullets, "1." for ordered lists.
- For diagrams/sketches you cannot transcribe, leave AT MOST ONE short marker "(см. рис.)" — never repeat it.
- If you cannot read part of the page (smudged, cut off), write "[нечитаемо]" inline — do NOT invent content.
- Use formal Russian academic terminology (Задача, Решение, Часть, Найдите, Покажите, что …).
- Do NOT translate identifiers, units, code, or proper names (BJT, MOSFET, V_T, …).
`;

const VERIFY_PROMPT = `Ты — рецензент перевода. Получишь изображение исходной страницы и черновой Russian Markdown.
Проверь:
1. Все ли формулы со страницы попали в Markdown?
2. Все ли подвопросы/пункты (а)(б)(в)(г)(д)(е) присутствуют?
3. Нет ли пропущенных абзацев?
4. Сохранён ли порядок элементов страницы: текст, таблицы, рисунки, графики, формулы?

ВЕРНИ строго JSON: {"ok":boolean,"gaps":["короткое описание пропуска", ...]}.
ok=true когда покрытие полное.`;

const PDFJS_RECONCILE_PROMPT = `Ты — дешёвый сверщик OCR/перевода с pdfjs-текстом.
Получишь:
1. pdfjs-текст исходной страницы (может быть с ошибками порядка чтения),
2. черновой русский Markdown.

Проверь, нет ли очевидно пропущенных терминов, формул, номеров задач, пунктов, строк таблиц.
ВЕРНИ строго JSON: {"ok":boolean,"gaps":["что пропущено или подозрительно", ...]}.
Не придирайся к нормальному перефразированию и переводу.`;

const FINAL_WATCHDOG_PROMPT = `Ты — финальный дешёвый watchdog академического перевода.
Проверь русский Markdown страницы перед сборкой DOCX:
- нет ли HTML-тегов;
- все формулы обёрнуты в $...$ или $$...$$;
- русский академический стиль, без китайских/японских символов;
- не потеряны очевидные элементы из layout blueprint;
- таблицы Markdown корректны.

ВЕРНИ строго JSON: {"ok":boolean,"issues":["короткое описание", ...]}.
ok=true если можно собирать DOCX.`;

const MATH_AUDIT_PROMPT = `Ты — строгий проверяющий математических формул.
Получишь изображение страницы и черновой Markdown.
Найди ВСЕ математические выражения на странице, которые в Markdown НЕ обёрнуты в $...$ или $$...$$
(например: y[n] = ..., x_1[n], H{x[n]}, \\delta[n-2], \\mathcal{H}, max(...), и подобные).

ВЕРНИ строго JSON: {"ok":boolean,"unwrapped":["цитата неправильного фрагмента", ...]}.
ok=true когда все формулы правильно обёрнуты.`;

async function verifyPage(
  page: { imageDataUrl: string; index: number },
  draft: string,
  opts: PlannerOpts,
  prompt: string,
  fieldKey: "gaps" | "unwrapped" | "issues",
): Promise<{ ok: boolean; issues: string[] }> {
  try {
    const visionUrl = await downsampleDataUrl(page.imageDataUrl, { maxDim: 1400 });
    const out = await chat({
      apiKey: opts.apiKey,
      model: economyVisionModel(opts.model),
      temperature: 0,
      maxTokens: 1500,
      signal: opts.signal,
      messages: [
        { role: "system", content: prompt },
        {
          role: "user",
          content: [
            { type: "text", text: `Стр. ${page.index + 1}. Черновой Markdown:\n\n${draft.slice(0, 8000)}` },
            { type: "image_url", image_url: { url: visionUrl } },
          ],
        },
      ],
    });
    const cleaned = stripCodeFences(out);
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (!m) return { ok: true, issues: [] };
    const parsed = JSON.parse(m[0]);
    const issues = Array.isArray(parsed[fieldKey]) ? parsed[fieldKey].slice(0, 10) : [];
    return { ok: parsed.ok !== false && issues.length === 0, issues };
  } catch {
    return { ok: true, issues: [] };
  }
}

async function reconcilePdfjs(
  page: { text: string; index: number },
  draft: string,
  opts: PlannerOpts,
): Promise<{ ok: boolean; issues: string[] }> {
  const hint = page.text.replace(/\s+/g, " ").trim();
  if (hint.length < 20) return { ok: true, issues: [] };
  try {
    const out = await chat({
      apiKey: opts.apiKey,
      model: economyVisionModel(opts.model),
      temperature: 0,
      maxTokens: 1000,
      signal: opts.signal,
      messages: [
        { role: "system", content: PDFJS_RECONCILE_PROMPT },
        { role: "user", content: `Стр. ${page.index + 1}\n\npdfjs-текст:\n${hint.slice(0, 9000)}\n\nЧерновой Markdown:\n${draft.slice(0, 9000)}` },
      ],
    });
    const cleaned = stripCodeFences(out);
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (!m) return { ok: true, issues: [] };
    const parsed = JSON.parse(m[0]);
    const issues = Array.isArray(parsed.gaps) ? parsed.gaps.slice(0, 10) : [];
    return { ok: parsed.ok !== false && issues.length === 0, issues };
  } catch {
    return { ok: true, issues: [] };
  }
}

function cleanMarkdown(raw: string): string {
  return wrapOrphanLatex(normalizeMath(normalizeMathDelims(sanitizeHtml(stripCodeFences(raw)))));
}

async function finalWatchdog(
  page: { index: number },
  draft: string,
  opts: PlannerOpts,
  blueprintHint: string,
): Promise<{ ok: boolean; issues: string[] }> {
  try {
    const out = await chat({
      apiKey: opts.apiKey,
      model: economyVisionModel(opts.model),
      temperature: 0,
      maxTokens: 1000,
      signal: opts.signal,
      messages: [
        { role: "system", content: FINAL_WATCHDOG_PROMPT },
        { role: "user", content: `Стр. ${page.index + 1}\n${blueprintHint}\n\nMarkdown:\n${draft.slice(0, 10000)}` },
      ],
    });
    const m = stripCodeFences(out).match(/\{[\s\S]*\}/);
    if (!m) return { ok: true, issues: [] };
    const parsed = JSON.parse(m[0]);
    const issues = Array.isArray(parsed.issues) ? parsed.issues.slice(0, 10) : [];
    return { ok: parsed.ok !== false && issues.length === 0, issues };
  } catch {
    return { ok: true, issues: [] };
  }
}

/* ─── Per-page translation ─────────────────────── */

async function translateDocPage(
  page: {
    text: string;
    imageDataUrl: string;
    index: number;
    images?: { dataUrl: string; y: number; w: number; h: number }[];
  },
  opts: PlannerOpts,
  glossary?: Glossary,
  blueprint?: DocumentBlueprint | null,
): Promise<DocBlock[]> {
  // Strategy: page IMAGE is always the primary OCR input.
  // pdfjs-extracted text is passed as a supplementary hint (may help on
  // dense printed pages, ignored when the image disagrees).

  const sysPrompt =
    VISION_OCR_PROMPT(TARGET_LANG) +
    dspGlossaryPrompt() +
    (glossary && glossary.size ? glossaryPrompt(glossary) : "");

  const visionUrl = page.imageDataUrl
    ? await downsampleDataUrl(page.imageDataUrl, { maxDim: 1800 })
    : "";

  const hint = page.text.replace(/\s+/g, " ").trim().slice(0, 8000);
  const hasHint = hint.length > 20;

  const layoutHint = pageBlueprintHint(blueprint?.pages?.[page.index]);

  const intro = `Page ${page.index + 1}. Read the attached image as your PRIMARY source — this is the authoritative version of the page (printed, handwritten, mixed, or scanned). Transcribe everything you see and translate to ${TARGET_LANG} as Markdown per the rules. Preserve the layout order from the blueprint: if the page contains a table, chart, figure, formula block, or diagram, leave its natural marker/position in the Markdown.`;

  const hintBlock = hasHint
    ? `\n\n---\nMachine-extracted text hint (from pdfjs/docx layout, may have OCR errors, missing handwriting/figures, or wrong reading order — use ONLY as a tie-breaker, trust the image when they disagree):\n\n${hint}`
    : "";

  const userContent: Array<
    | { type: "text"; text: string }
    | { type: "image_url"; image_url: { url: string } }
  > = [
    { type: "text", text: intro + layoutHint + hintBlock },
    ...(visionUrl
      ? [{ type: "image_url" as const, image_url: { url: visionUrl } }]
      : []),
  ];

  opts.onLog?.(
    `Стр. ${page.index + 1}: page OCR ${economyVisionModel(opts.model)}${hasHint ? " + pdfjs сверка" : ""}${visionUrl ? "" : " (нет изображения)"}`,
  );

  const out = await chat({
    apiKey: opts.apiKey,
    model: economyVisionModel(opts.model),
    temperature: 0.2,
    maxTokens: 4096,
    signal: opts.signal,
    messages: [
      { role: "system", content: sysPrompt },
      { role: "user", content: userContent },
    ],
  });
  let raw = out;
  // ─── Pass 2: coverage verify ──────────────────────────────────────────
  // ─── Pass 3: math audit ──────────────────────────────────────────────
  if (page.imageDataUrl) {
    for (const pass of [
      { prompt: VERIFY_PROMPT, field: "gaps" as const, label: "verify" },
      { prompt: MATH_AUDIT_PROMPT, field: "unwrapped" as const, label: "math-audit" },
    ]) {
      try {
        const check = await verifyPage({ imageDataUrl: page.imageDataUrl, index: page.index }, raw, opts, pass.prompt, pass.field);
        if (!check.ok && check.issues.length) {
          opts.onLog?.(`Стр. ${page.index + 1}: ${pass.label} нашёл ${check.issues.length} проблем — повторный перевод…`);
          const fixOut = await chat({
            apiKey: opts.apiKey,
            model: economyVisionModel(opts.model),
            temperature: 0.1,
            maxTokens: 4096,
            signal: opts.signal,
            messages: [
              { role: "system", content: sysPrompt },
              { role: "user", content: userContent },
              { role: "assistant", content: raw },
              {
                role: "user",
                content: `Исправь следующие проблемы и верни ПОЛНУЮ обновлённую версию страницы как Markdown (без комментариев). Используй layout blueprint и pdfjs hint как вспомогательные источники, но изображение — главное:\n${check.issues.map((s, i) => `${i + 1}. ${s}`).join("\n")}`,
              },
            ],
          });
          raw = fixOut;
        }
      } catch { /* ignore verify errors */ }
    }
  }

  const pdfjsCheck = await reconcilePdfjs(page, raw, opts);
  if (!pdfjsCheck.ok && pdfjsCheck.issues.length) {
    opts.onLog?.(`Стр. ${page.index + 1}: pdfjs-сверка нашла ${pdfjsCheck.issues.length} проблем — корректировка…`);
    raw = await chat({
      apiKey: opts.apiKey,
      model: economyVisionModel(opts.model),
      temperature: 0.1,
      maxTokens: 4096,
      signal: opts.signal,
      messages: [
        { role: "system", content: sysPrompt },
        { role: "user", content: userContent },
        { role: "assistant", content: raw },
        { role: "user", content: `Сверка с pdfjs нашла проблемы. Верни ПОЛНУЮ исправленную страницу Markdown, ничего не комментируй:\n${pdfjsCheck.issues.map((s, i) => `${i + 1}. ${s}`).join("\n")}` },
      ],
    });
  }

  let cleaned = cleanMarkdown(raw);
  const watchdog = await finalWatchdog(page, cleaned, opts, layoutHint);
  if (!watchdog.ok && watchdog.issues.length) {
    opts.onLog?.(`Стр. ${page.index + 1}: watchdog нашёл ${watchdog.issues.length} проблем — финальная правка…`);
    const fixed = await chat({
      apiKey: opts.apiKey,
      model: economyVisionModel(opts.model),
      temperature: 0,
      maxTokens: 4096,
      signal: opts.signal,
      messages: [
        { role: "system", content: sysPrompt },
        { role: "user", content: userContent },
        { role: "assistant", content: cleaned },
        { role: "user", content: `Исправь только эти проблемы и верни ПОЛНУЮ страницу Markdown:\n${watchdog.issues.map((s, i) => `${i + 1}. ${s}`).join("\n")}` },
      ],
    });
    cleaned = cleanMarkdown(fixed);
  }

  const blocks = parseMarkdownToBlocks(cleaned);
  // Post-pass: substitute any remaining English DSP terms.
  return blocks.map((b) => {
    if (b.type === "para" || b.type === "h1" || b.type === "h2" || b.type === "h3") {
      return { ...b, text: polishRu(applyGlossaryPost(b.text)) };
    }
    if (b.type === "list") {
      return { ...b, items: b.items.map((it) => polishRu(applyGlossaryPost(it))) };
    }
    return b;
  });
}

/* ─── Full document planning ───────────────────── */

export async function planDoc(
  extracted: ExtractedDoc,
  opts: PlannerOpts,
): Promise<DocPlan> {
  const allBlocks: DocBlock[] = [];
  let title: string | undefined;
  const errors: string[] = [];
  const glossary: Glossary = new Map();

  let done = 0;
  const total = extracted.pages.length;
  const pagePairs: Array<Array<[string, string]>> = new Array(total).fill([]);

  const blueprint = await buildDocumentBlueprint(extracted, opts);

  const results = await mapWithConcurrency(
    extracted.pages,
    Math.max(1, opts.concurrency ?? 3),
    async (page, i) => {
      const blocks = await translateDocPage(page, opts, glossary, blueprint);
      done++;
      opts.onProgress?.(done, total);
      opts.onLog?.(`Стр. ${i + 1}/${total} переведена`);
      const tt = blocks
        .map((b) =>
          b.type === "para" || b.type === "h1" || b.type === "h2" || b.type === "h3"
            ? b.text
            : b.type === "list"
              ? b.items.join(" ")
              : "",
        )
        .join("\n");
      pagePairs[i] = harvestPairs(page.text, tt);
      return blocks;
    },
    {
      signal: opts.signal,
      onBatchSettled: (start, end) => {
        for (let i = start; i <= end; i++) mergeGlossary(glossary, pagePairs[i]);
        if (glossary.size > 0)
          opts.onLog?.(`Глоссарий: ${glossary.size} терминов`);
      },
    },
  );

  for (let i = 0; i < extracted.pages.length; i++) {
    const page = extracted.pages[i];
    const r = results[i];
    if (r.ok) {
      const blocks = r.value;
      if (!title && blocks.length > 0 && blocks[0].type === "h1") {
        const h1 = blocks.shift() as DocBlock;
        if (h1.type === "h1") title = h1.text;
      }
      if (blocks.length === 0) {
        if (page.text.trim())
          allBlocks.push({ type: "para", text: page.text.trim() });
      } else {
        allBlocks.push(...blocks);
      }
    } else {
      const msg = r.error.message;
      errors.push(`Страница ${i + 1}: ${msg}`);
      opts.onLog?.(`Ошибка на странице ${i + 1}: ${msg.slice(0, 120)}`);
      allBlocks.push({
        type: "para",
        text: `⚠ Страница ${i + 1}: не удалось перевести (${msg}). Исходный текст ниже.`,
      });
      if (page.text.trim())
        allBlocks.push({ type: "para", text: page.text.trim() });
    }

    // Figure extraction: ask vision model for bboxes, crop them.
    let figures: Array<{ dataUrl: string; caption?: string }> = [];
    if (page.imageDataUrl) {
      try {
        const bboxes = await detectFigures(page.imageDataUrl, {
          apiKey: opts.apiKey,
          model: opts.model,
          signal: opts.signal,
        });
        if (bboxes.length > 0) {
          const cropped = await cropFigures(page.imageDataUrl, bboxes);
          figures = cropped.map((c) => ({
            dataUrl: c.dataUrl,
            caption: c.bbox.caption,
          }));
          opts.onLog?.(`Стр. ${i + 1}: vision нашёл ${figures.length} рисунк(ов)`);
        }
      } catch (e) {
        opts.onLog?.(`Стр. ${i + 1}: детекция рисунков пропущена (${String(e).slice(0, 60)})`);
      }
    }

    // Fallback to pdfjs-embedded rasters ONLY if vision returned nothing.
    if (figures.length === 0 && page.images && page.images.length) {
      const pageW = page.width || 1;
      const pageH = page.height || 1;
      const realFigs = page.images.filter((im) => {
        const coverage = (im.w * im.h) / (pageW * pageH);
        return coverage > 0.005 && coverage < 0.95;
      });
      figures = realFigs.map((f) => ({ dataUrl: f.dataUrl }));
      if (figures.length)
        opts.onLog?.(`Стр. ${i + 1}: использую ${figures.length} pdfjs-картин(ок)`);
    }

    for (let k = 0; k < figures.length; k++) {
      const f = figures[k];
      allBlocks.push({
        type: "figure",
        imageDataUrl: f.dataUrl,
        caption:
          f.caption
          || (figures.length === 1
            ? `Рис. ${i + 1}`
            : `Рис. ${i + 1}.${k + 1}`),
      });
    }
    // NO whole-page fallback. If the model finds no figures and pdfjs has none,
    // the page is treated as text-only.
  }

  if (errors.length === extracted.pages.length) {
    throw new Error(
      `Перевод не удался ни на одной странице: ${errors[0]}`,
    );
  }

  return { title, blocks: allBlocks };
}
