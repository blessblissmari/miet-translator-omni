import { chat, parseJsonLoose } from "./mimo";
import { downsampleDataUrl } from "./imageOps";
import type { ExtractedDoc } from "./types";
import type { PlannerOpts } from "./plannerShared";

export interface LayoutElement {
  kind: "title" | "heading" | "paragraph" | "formula" | "table" | "figure" | "chart" | "diagram" | "image" | "footer" | "other";
  order: number;
  description?: string;
  bbox?: { x: number; y: number; w: number; h: number };
}

export interface PageBlueprint {
  pageIndex: number;
  pageType?: "cover" | "toc" | "text" | "exercise" | "solution" | "slide" | "scan" | "mixed";
  readingOrder?: string[];
  elements: LayoutElement[];
  notes?: string;
}

export interface DocumentBlueprint {
  documentKind?: string;
  language?: string;
  academicDomain?: string;
  translationStrategy?: string;
  pages: PageBlueprint[];
}

const GLOBAL_LAYOUT_PROMPT = `You are the first pass of an academic document translation pipeline.

Goal: analyze the WHOLE source document before page-by-page translation. Build a layout/OCR blueprint: what each page contains, reading order, and where figures/tables/formulas should appear.

Return ONLY strict JSON:
{
  "documentKind": "...",
  "language": "...",
  "academicDomain": "...",
  "translationStrategy": "...",
  "pages": [
    {
      "pageIndex": 0,
      "pageType": "cover|toc|text|exercise|solution|slide|scan|mixed",
      "readingOrder": ["short ordered notes"],
      "elements": [
        {"kind":"heading|paragraph|formula|table|figure|chart|diagram|image|footer|other", "order": 1, "description":"...", "bbox":{"x":0.1,"y":0.2,"w":0.5,"h":0.1}}
      ],
      "notes": "risks, handwriting, columns, important terms"
    }
  ]
}

Rules:
- Treat page images as authoritative and pdfjs text only as a hint.
- Include every page in pages[].
- For figures/charts/diagrams/images, bbox is normalized 0..1 from top-left when visually inferable.
- For formulas, describe formula groups and their reading order, but do not translate yet.
- Translation target later is formal Russian academic style.
- Do not include commentary outside JSON.`;

export function bestLayoutModel(model: string): string {
  return model || "mimo-v2.5";
}

export function economyVisionModel(model: string): string {
  return model === "mimo-v2.5" ? "mimo-v2-omni" : (model || "mimo-v2-omni");
}

export function pageBlueprintHint(bp?: PageBlueprint | null): string {
  if (!bp) return "";
  const elements = (bp.elements || [])
    .slice()
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
    .map((e) => {
      const box = e.bbox ? ` bbox=${e.bbox.x.toFixed(2)},${e.bbox.y.toFixed(2)},${e.bbox.w.toFixed(2)},${e.bbox.h.toFixed(2)}` : "";
      return `- ${e.order}. ${e.kind}${box}: ${e.description || ""}`;
    })
    .join("\n");
  const order = bp.readingOrder?.length ? `\nReading order:\n${bp.readingOrder.map((x, i) => `${i + 1}. ${x}`).join("\n")}` : "";
  return `\n\n---\nLayout blueprint for this page (from the whole-document first pass):\npageType: ${bp.pageType || "unknown"}${order}\nElements:\n${elements || "- none"}\nNotes: ${bp.notes || ""}`;
}

export async function buildDocumentBlueprint(
  extracted: ExtractedDoc,
  opts: PlannerOpts,
): Promise<DocumentBlueprint | null> {
  if (!extracted.pages.length) return null;

  const textManifest = extracted.pages.map((p) => {
    const lines = (p.lines || [])
      .slice(0, 80)
      .map((ln) => `[x=${ln.x.toFixed(1)} y=${ln.y.toFixed(1)}] ${ln.text}`)
      .join("\n");
    const text = lines || p.text.slice(0, 5000);
    return `page ${p.index + 1}/${extracted.pages.length}\nsize=${Math.round(p.width)}x${Math.round(p.height)}\n${text}`;
  }).join("\n\n---\n\n");

  const content: Array<
    | { type: "text"; text: string }
    | { type: "image_url"; image_url: { url: string } }
  > = [
    { type: "text", text: `Document has ${extracted.pages.length} pages. Here is pdfjs text/line layout for every page:\n\n${textManifest.slice(0, 120_000)}` },
  ];

  try {
    for (const page of extracted.pages) {
      if (!page.imageDataUrl) continue;
      content.push({ type: "text", text: `page ${page.index + 1} image:` });
      content.push({
        type: "image_url",
        image_url: { url: await downsampleDataUrl(page.imageDataUrl, { maxDim: 640, quality: 0.58 }) },
      });
    }

    opts.onLog?.(`Глобальный OCR/layout: анализ ${extracted.pages.length} стр. через ${bestLayoutModel(opts.model)}…`);
    const out = await chat({
      apiKey: opts.apiKey,
      model: bestLayoutModel(opts.model),
      temperature: 0,
      maxTokens: Math.min(12000, Math.max(3000, extracted.pages.length * 450)),
      responseJson: true,
      signal: opts.signal,
      retries: 2,
      messages: [
        { role: "system", content: GLOBAL_LAYOUT_PROMPT },
        { role: "user", content },
      ],
    });
    const parsed = parseJsonLoose<DocumentBlueprint>(out);
    if (!Array.isArray(parsed.pages)) return null;
    opts.onLog?.(`Глобальный OCR/layout готов: ${parsed.pages.length} стр.`);
    return normalizeBlueprint(parsed, extracted.pages.length);
  } catch (e) {
    opts.onLog?.(`Глобальный OCR/layout пропущен: ${(e as Error).message.slice(0, 120)}`);
    return null;
  }
}

function normalizeBlueprint(bp: DocumentBlueprint, pageCount: number): DocumentBlueprint {
  const byIndex = new Map<number, PageBlueprint>();
  for (const p of bp.pages || []) {
    const pageIndex = Number.isFinite(p.pageIndex) ? p.pageIndex : byIndex.size;
    byIndex.set(pageIndex, {
      ...p,
      pageIndex,
      elements: Array.isArray(p.elements) ? p.elements : [],
    });
  }
  const pages: PageBlueprint[] = [];
  for (let i = 0; i < pageCount; i++) {
    pages.push(byIndex.get(i) || { pageIndex: i, pageType: "mixed", elements: [] });
  }
  return { ...bp, pages };
}
