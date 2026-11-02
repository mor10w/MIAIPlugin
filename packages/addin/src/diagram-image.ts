import { diagramToSvg, type DiagramResult } from "@miai/shared";

export interface PngImage {
  /** Base64 uden data-url-præfiks (klar til Office.js' *FromBase64-metoder). */
  base64: string;
  width: number;
  height: number;
}

/** Renderer diagrammet som PNG (2x for skarphed) og returnerer base64. */
export async function diagramToPng(d: DiagramResult, scale = 2): Promise<PngImage> {
  const svg = diagramToSvg(d);
  const m = /width="([\d.]+)" height="([\d.]+)"/.exec(svg);
  const w = Math.round((m ? Number(m[1]) : 600) * scale);
  const h = Math.round((m ? Number(m[2]) : 400) * scale);
  const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }));
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error("Kunne ikke render diagrammet."));
      i.src = url;
    });
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx2d = canvas.getContext("2d");
    if (!ctx2d) throw new Error("Canvas understøttes ikke i dette webview.");
    ctx2d.drawImage(img, 0, 0, w, h);
    const dataUrl = canvas.toDataURL("image/png");
    return { base64: dataUrl.slice("data:image/png;base64,".length), width: w, height: h };
  } finally {
    URL.revokeObjectURL(url);
  }
}
