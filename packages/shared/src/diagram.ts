import { z } from "zod";

/**
 * Diagram som struktureret output (ADR-0009): knuder + kanter.
 * Bevidst simpelt JSON i stedet for Mermaid: lokale modeller producerer det mere
 * pålideligt, og add-in'et kan uden parser både rendere SVG (Word) og oprette
 * native shapes i PowerPoint.
 */
export const DiagramResult = z.object({
  title: z.string().max(120).optional(),
  nodes: z.array(z.object({ id: z.string().min(1).max(40), label: z.string().min(1).max(80) })).min(1).max(24),
  edges: z.array(z.object({ from: z.string().min(1).max(40), to: z.string().min(1).max(40) })).max(60).optional(),
});

export type DiagramResult = z.infer<typeof DiagramResult>;

export interface LaidNode {
  id: string;
  label: string;
  /** Position og størrelse i layoutenheder (ca. punkter ved 96 dpi). */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface DiagramLayout {
  nodes: LaidNode[];
  width: number;
  height: number;
}

const MARGIN = 12;
const X_GAP = 64; // mellem lag (horisontalt)
const Y_GAP = 28; // mellem knuder i samme lag (vertikalt)
const NODE_H = 40;
const LINE_H = 18;

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Deler et label i op til 3 linjer (grædsk, maksimalt ~28 tegn/linje). */
function wrap(label: string): string[] {
  const words = label.split(/\s+/);
  const lines: string[] = [];
  let cur = "";
  for (const w of words) {
    if ((cur ? `${cur} ${w}` : w).length <= 28 || !cur) cur = cur ? `${cur} ${w}` : w;
    else {
      lines.push(cur);
      cur = w;
    }
  }
  if (cur) lines.push(cur);
  return lines.slice(0, 3);
}

/**
 * Deterministisk layout: venstre-modet-højre, lag efter længste sti fra rødder
 * (khan-topologisk sortering; knuder i cykler placeres i det sidste lag).
 */
export function layoutDiagram(d: DiagramResult): DiagramLayout {
  const nodes = d.nodes;
  const ids = new Set(nodes.map((n) => n.id));
  const edges = (d.edges ?? []).filter((e) => ids.has(e.from) && ids.has(e.to));

  const indeg = new Map<string, number>(nodes.map((n) => [n.id, 0]));
  const out = new Map<string, string[]>();
  for (const e of edges) {
    out.set(e.from, [...(out.get(e.from) ?? []), e.to]);
    indeg.set(e.to, (indeg.get(e.to) ?? 0) + 1);
  }

  // Lag = længste sti fra en rod. Iterativt (ingen rekursionsdybde).
  const layer = new Map<string, number>();
  const queue: string[] = [];
  for (const n of nodes) if ((indeg.get(n.id) ?? 0) === 0) queue.push(n.id);
  const processed = new Set<string>();
  while (queue.length) {
    const id = queue.shift()!;
    if (processed.has(id)) continue;
    processed.add(id);
    const l = layer.get(id) ?? 0;
    for (const to of out.get(id) ?? []) {
      layer.set(to, Math.max(layer.get(to) ?? 0, l + 1));
      indeg.set(to, (indeg.get(to) ?? 0) - 1);
      if ((indeg.get(to) ?? 0) === 0) queue.push(to);
    }
  }
  // Cyklusknuder: placer i det dybeste lag + 1.
  let maxLayer = 0;
  for (const l of layer.values()) maxLayer = Math.max(maxLayer, l);
  for (const n of nodes) if (!processed.has(n.id)) layer.set(n.id, maxLayer + 1);

  // Grupper i lag (bevarer input-rækkefølge = deterministisk).
  const layers: string[][] = [];
  for (const n of nodes) {
    const l = layer.get(n.id) ?? 0;
    (layers[l] ??= []).push(n.id);
  }

  const size = new Map<string, { w: number; lines: string[]; h: number }>();
  for (const n of nodes) {
    const lines = wrap(n.label);
    const w = Math.max(90, Math.min(280, ...lines.map((l) => l.length * 7 + 36), n.label.length * 7 + 36));
    const h = Math.max(NODE_H, lines.length * LINE_H + 14);
    size.set(n.id, { w, lines, h });
  }

  const laid: LaidNode[] = [];
  let xCursor = MARGIN;
  let maxStackH = 0;
  for (const idsInLayer of layers) {
    if (!idsInLayer || !idsInLayer.length) continue; // tomme lag (fx. cyklus) springes over
    const maxW = Math.max(...idsInLayer.map((id) => size.get(id)!.w));
    let y = MARGIN; // alle lag starter øverst → rene rækker/træer, ikke en trappe
    for (const id of idsInLayer) {
      const s = size.get(id)!;
      laid.push({ id, label: nodes.find((n) => n.id === id)!.label, x: xCursor + (maxW - s.w) / 2, y, w: s.w, h: s.h });
      y += s.h + Y_GAP;
    }
    xCursor += maxW + X_GAP;
    maxStackH = Math.max(maxStackH, y - Y_GAP);
  }

  return { nodes: laid, width: Math.max(xCursor - X_GAP + MARGIN, 10), height: Math.max(maxStackH + MARGIN, 10) };
}

/**
 * Renderer diagrammet som selvstændigt SVG (runden rektangler, tekst og pile).
 * Brugt til forhåndsgreb i taskpane'et og som kilde til PNG (Word).
 */
export function diagramToSvg(d: DiagramResult): string {
  const L = layoutDiagram(d);
  const pos = new Map(L.nodes.map((n) => [n.id, n]));
  const parts: string[] = [];

  // Kanter først (så knuder tegnes ovenpå). Pile: linje + trekant.
  for (const e of d.edges ?? []) {
    const a = pos.get(e.from);
    const b = pos.get(e.to);
    if (!a || !b) continue;
    const x1 = a.x + a.w, y1 = a.y + a.h / 2;
    const x2 = b.x, y2 = b.y + b.h / 2;
    const mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
    // Let buet sti giver pænere krydsning end rette linjer.
    parts.push(`<path d="M ${x1} ${y1} C ${mx} ${y1}, ${mx} ${y2}, ${x2} ${y2}" fill="none" stroke="#6b7280" stroke-width="1.5"/>`);
    const ang = Math.atan2(y2 - my, x2 - mx);
    parts.push(
      `<path d="M ${x2} ${y2} l ${-9 * Math.cos(ang - 0.4)} ${-9 * Math.sin(ang - 0.4)} l ${-9 * Math.cos(ang + 0.4)} ${-9 * Math.sin(ang + 0.4)} z" fill="#6b7280"/>`,
    );
  }

  for (const n of L.nodes) {
    const lines = wrap(n.label);
    parts.push(
      `<rect x="${n.x}" y="${n.y}" width="${n.w}" height="${n.h}" rx="8" fill="#eff6ff" stroke="#3b82f6"/>`,
      `<text x="${n.x + n.w / 2}" y="${n.y + n.h / 2 - ((lines.length - 1) * LINE_H) / 2 + 5}" text-anchor="middle" font-family="-apple-system, 'Segoe UI', sans-serif" font-size="13" fill="#111827">${lines
        .map((l, i) => (i === 0 ? escapeXml(l) : `tspan x="${n.x + n.w / 2}" dy="18">${escapeXml(l)}`))
        .join("")}</text>`,
    );
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${L.width}" height="${L.height}" viewBox="0 0 ${L.width} ${L.height}" style="max-width:100%;height:auto">${parts.join("")}</svg>`;
}
