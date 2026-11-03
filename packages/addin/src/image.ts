import { diagramToSvg, type DiagramResult } from "@miai/shared";
import { mermaidToSvg } from "./mermaid.js";
import type { GraphicBlock } from "./svg.js";
import { sanitizeSvg } from "./svg.js";

export interface PngImage {
  /** Base64 uden data-url-præfiks (klar til Office.js' *FromBase64-metoder). */
  base64: string;
  width: number;
  height: number;
  /** Den skala, der faktisk blev brugt (kan være < 2 for store figurer) – giver naturlig størrelse = width / scale. */
  scale: number;
}

/** Naturlig størrelse i punkter (1 px = 0,75 pt), højst `maxWidth` bred. */
export function pointSize(png: PngImage, maxWidth = 450): { width: number; height: number } {
  const width = Math.min(maxWidth, (png.width / png.scale) * 0.75);
  return { width, height: width * (png.height / png.width) };
}

/** Læser dimensioner fra SVG (viewBox foretrækkes – mermaid skriver width="100%"). */
export function svgSize(svg: string): { w: number; h: number } {
  // viewBox kan have negative/decimale tal og komma-separering (Mermaid skriver fx "-8 -8 512.5 300").
  const vb = /viewBox="\s*-?[\d.]+[\s,]+-?[\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)\s*"/.exec(svg);
  if (vb && Number(vb[1]) > 0 && Number(vb[2]) > 0) return { w: Number(vb[1]), h: Number(vb[2]) };
  const wh = /width="([\d.]+)"\s+height="([\d.]+)"/.exec(svg);
  return wh ? { w: Number(wh[1]), h: Number(wh[2]) } : { w: 600, h: 400 };
}

/** Konverterer en SVG-streng til PNG og returnerer base64. Maks. 2048 px pr. side. */
export async function svgToPng(svg: string, scale = 2): Promise<PngImage> {
  const { w, h } = svgSize(svg);
  // Cap: store viewBox'er (fx. fra modellen) må ikke rasterisere til kæmpe canvas'er.
  const s = Math.min(scale, 2048 / Math.max(w, h));
  const W = Math.max(1, Math.round(w * s));
  const H = Math.max(1, Math.round(h * s));
  const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }));
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error("Kunne ikke render diagrammet."));
      i.src = url;
    });
    const canvas = document.createElement("canvas");
    canvas.width = W;
    canvas.height = H;
    const ctx2d = canvas.getContext("2d");
    if (!ctx2d) throw new Error("Canvas understøttes ikke i dette webview.");
    ctx2d.drawImage(img, 0, 0, W, H);
    const dataUrl = canvas.toDataURL("image/png");
    return { base64: dataUrl.slice("data:image/png;base64,".length), width: W, height: H, scale: s };
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Renderer diagrammet som PNG (2x for skarphed) og returnerer base64. */
export async function diagramToPng(d: DiagramResult, scale = 2): Promise<PngImage> {
  return svgToPng(diagramToSvg(d), scale);
}

/** Renderer Mermaid-kode som PNG og returnerer base64. */
export async function mermaidToPng(code: string, scale = 2): Promise<PngImage> {
  return svgToPng(await mermaidToSvg(code), scale);
}

/** Konverterer den første ```mermaid- eller ```svg-blok i teksten til PNG (ADR-0009). */
export async function graphicToPng(block: GraphicBlock, scale = 2): Promise<PngImage> {
  if (block.kind === "mermaid") return mermaidToPng(block.code, scale);
  return svgToPng(sanitizeSvg(block.code), scale); // SVG er modeloutput og skal renses først
}
