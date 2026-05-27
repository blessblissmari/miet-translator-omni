import type { ExtractedDoc, ExtractedImage, ExtractedPage } from "./types";
import { renderAsync } from "docx-preview";

type ExtractedLine = NonNullable<ExtractedPage["lines"]>[number];

export async function extractDocxRendered(blob: Blob): Promise<ExtractedDoc> {
  const host = document.createElement("div");
  host.style.cssText = "position:fixed;left:-100000px;top:0;width:1200px;background:white;visibility:hidden;pointer-events:none;";
  document.body.appendChild(host);
  try {
    await renderAsync(blob, host, undefined, {
      className: "docx-source",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
      experimental: true,
    });

    const sections = Array.from(host.querySelectorAll<HTMLElement>("section.docx-source"));
    const pages = sections.length ? sections : [host];
    const extractedPages: ExtractedPage[] = [];

    for (let i = 0; i < pages.length; i++) {
      const pageEl = pages[i];
      const rect = pageEl.getBoundingClientRect();
      const width = Math.max(1, Math.round(rect.width || pageEl.scrollWidth || 794));
      const height = Math.max(1, Math.round(rect.height || pageEl.scrollHeight || 1123));
      const lines = extractLines(pageEl, rect);
      const images = await extractImages(pageEl, rect);
      const text = lines.length
        ? lines.map((l) => l.text).join("\n")
        : (pageEl.textContent || "").replace(/\n{3,}/g, "\n\n").trim();
      const imageDataUrl = await renderDomPageFallback(text, width, height, images);
      extractedPages.push({ index: i, text, imageDataUrl, width, height, lines, images });
    }

    return { pages: extractedPages, meta: {} };
  } finally {
    host.remove();
  }
}

function extractLines(pageEl: HTMLElement, pageRect: DOMRect): ExtractedLine[] {
  const out: ExtractedLine[] = [];
  const walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const text = node.textContent?.replace(/\s+/g, " ").trim() || "";
      return text ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
    },
  });
  let node: Node | null;
  while ((node = walker.nextNode())) {
    const text = node.textContent?.replace(/\s+/g, " ").trim() || "";
    if (!text) continue;
    const range = document.createRange();
    range.selectNodeContents(node);
    const rects = Array.from(range.getClientRects()).filter((r) => r.width > 1 && r.height > 1);
    range.detach();
    if (!rects.length) continue;
    const parent = node.parentElement as HTMLElement | null;
    const style = parent ? getComputedStyle(parent) : null;
    const fontSize = style ? parseFloat(style.fontSize || "14") : 14;
    for (const r of rects) {
      out.push({
        text,
        x: r.left - pageRect.left,
        y: r.top - pageRect.top,
        w: r.width,
        h: r.height,
        fontSize,
      });
    }
  }
  out.sort((a, b) => a.y - b.y || a.x - b.x);
  return mergeNearbyLineFragments(out);
}

function mergeNearbyLineFragments(lines: ExtractedLine[]): ExtractedLine[] {
  const merged: ExtractedLine[] = [];
  for (const ln of lines) {
    const last = merged[merged.length - 1];
    if (last && Math.abs(last.y - ln.y) < Math.max(3, last.h * 0.35) && ln.x >= last.x) {
      const gap = ln.x - (last.x + last.w);
      last.text += gap > last.fontSize * 0.3 ? ` ${ln.text}` : ln.text;
      const right = Math.max(last.x + last.w, ln.x + ln.w);
      last.w = right - last.x;
      last.h = Math.max(last.h, ln.h);
      last.fontSize = Math.max(last.fontSize, ln.fontSize);
    } else {
      merged.push({ ...ln });
    }
  }
  return merged.filter((l) => l.text.trim());
}

async function extractImages(pageEl: HTMLElement, pageRect: DOMRect): Promise<ExtractedImage[]> {
  const out: ExtractedImage[] = [];
  for (const img of Array.from(pageEl.querySelectorAll<HTMLImageElement>("img"))) {
    const rect = img.getBoundingClientRect();
    if (rect.width < 16 || rect.height < 16) continue;
    const dataUrl = await imageToDataUrl(img).catch(() => "");
    if (!dataUrl) continue;
    out.push({
      dataUrl,
      y: rect.top - pageRect.top,
      w: rect.width,
      h: rect.height,
    });
  }
  out.sort((a, b) => a.y - b.y);
  return out;
}

async function imageToDataUrl(img: HTMLImageElement): Promise<string> {
  const src = img.currentSrc || img.src;
  if (src.startsWith("data:")) return src;
  const res = await fetch(src);
  const blob = await res.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function renderDomPageFallback(text: string, w: number, h: number, images: ExtractedImage[]): Promise<string> {
  const maxW = Math.min(1600, Math.max(900, w));
  const scale = maxW / Math.max(1, w);
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(w * scale);
  canvas.height = Math.round(h * scale);
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#111111";
  ctx.font = `${Math.max(16, Math.round(16 * scale))}px Times New Roman`;
  let y = Math.round(50 * scale);
  for (const line of wrapText(ctx, text, canvas.width - Math.round(90 * scale))) {
    if (y > canvas.height - 30) break;
    ctx.fillText(line, Math.round(45 * scale), y);
    y += Math.round(24 * scale);
  }
  for (const im of images.slice(0, 12)) {
    try {
      const image = await loadImage(im.dataUrl);
      const drawX = Math.round(45 * scale);
      const drawY = Math.min(canvas.height - 120, Math.round((im.y || y / scale) * scale));
      const drawW = Math.min(canvas.width - 90 * scale, im.w * scale);
      const drawH = Math.min(canvas.height / 3, im.h * scale);
      ctx.drawImage(image, drawX, drawY, drawW, drawH);
    } catch { /* ignore */ }
  }
  return canvas.toDataURL("image/jpeg", 0.88);
}

function wrapText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const out: string[] = [];
  for (const para of text.split(/\n/)) {
    const words = para.split(/\s+/).filter(Boolean);
    let line = "";
    for (const word of words) {
      const next = line ? `${line} ${word}` : word;
      if (ctx.measureText(next).width > maxWidth && line) {
        out.push(line);
        line = word;
      } else {
        line = next;
      }
    }
    if (line) out.push(line);
  }
  return out;
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}
