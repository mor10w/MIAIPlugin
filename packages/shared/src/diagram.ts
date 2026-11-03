import { z } from "zod";

/**
 * Diagram som struktureret output (ADR-0009): knuder + kanter.
 * Bevidst simpelt JSON i stedet for Mermaid: lokale modeller producerer det mere
 * pålideligt, og add-in'et kan uden parser både rendere SVG (Word) og oprette
 * native shapes i PowerPoint.
 */
export const DiagramResult = z.object({
  title: z.string().max(120).optional(),
  nodes: z
    .array(z.object({ id: z.string().min(1).max(40), label: z.string().min(1).max(80) }))
    .min(1)
    .max(24),
  edges: z
    .array(z.object({ from: z.string().min(1).max(40), to: z.string().min(1).max(40) }))
    .max(60)
    .optional(),
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
  /** Søjlens (lagets) venstre kant og bredde – knuder centreres i søjlen. Bruges til kantruter. */
  colX: number;
  colW: number;
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

/** Deler et label i op til 3 linjer (grådig, maks. ~28 tegn/linje). */
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
  if (lines.length > 3) lines[2] = `${lines[2].slice(0, 27)}…`;
  return lines.slice(0, 3);
}

/**
 * Deterministisk layout: venstre mod højre. Cykler brydes først (DFS-tilbagekanter),
 * derefter lag efter længste sti fra rødderne (Kahns topologiske sortering).
 */
export function layoutDiagram(d: DiagramResult): DiagramLayout {
  const nodes = d.nodes;
  const ids = new Set(nodes.map((n) => n.id));
  const edges = (d.edges ?? []).filter((e) => ids.has(e.from) && ids.has(e.to));

  // 1) Bryd cykler: DFS i input-rækkefølge; kanter til en knude på stakken er tilbagekanter.
  //    De ignoreres ved lagdelingen (men tegnes stadig som returkanter i routeEdges).
  //    Uden dette ville en enkelt feedback-løkke (fx Rapport → Indsamling) fjerne alle rødder
  //    og lægge hele diagrammet i én søjle.
  const out = new Map<string, string[]>();
  for (const e of edges) out.set(e.from, [...(out.get(e.from) ?? []), e.to]);
  const state = new Map<string, 1 | 2>(); // 1 = på stakken, 2 = færdig
  const back = new Set<string>();
  for (const start of nodes) {
    if (state.has(start.id)) continue;
    // Iterativ DFS (ingen rekursionsdybde).
    const stack: { id: string; i: number }[] = [{ id: start.id, i: 0 }];
    state.set(start.id, 1);
    while (stack.length) {
      const top = stack[stack.length - 1];
      const next = (out.get(top.id) ?? [])[top.i++];
      if (next === undefined) {
        state.set(top.id, 2);
        stack.pop();
      } else if (state.get(next) === 1) {
        back.add(`${top.id}\u0000${next}`);
      } else if (!state.has(next)) {
        state.set(next, 1);
        stack.push({ id: next, i: 0 });
      }
    }
  }
  const dag = edges.filter((e) => !back.has(`${e.from}\u0000${e.to}`) && e.from !== e.to);

  // 2) Lag = længste sti fra en rod (Kahn på den cyklusfri graf).
  const indeg = new Map<string, number>(nodes.map((n) => [n.id, 0]));
  const succ = new Map<string, string[]>();
  for (const e of dag) {
    succ.set(e.from, [...(succ.get(e.from) ?? []), e.to]);
    indeg.set(e.to, (indeg.get(e.to) ?? 0) + 1);
  }
  const layer = new Map<string, number>();
  const queue = nodes.filter((n) => (indeg.get(n.id) ?? 0) === 0).map((n) => n.id);
  while (queue.length) {
    const id = queue.shift()!;
    const l = layer.get(id) ?? 0;
    for (const to of succ.get(id) ?? []) {
      layer.set(to, Math.max(layer.get(to) ?? 0, l + 1));
      indeg.set(to, (indeg.get(to) ?? 0) - 1);
      if ((indeg.get(to) ?? 0) === 0) queue.push(to);
    }
  }

  // Grupper i lag (bevarer input-rækkefølge = deterministisk).
  const layers: string[][] = [];
  for (const n of nodes) {
    const l = layer.get(n.id) ?? 0;
    (layers[l] ??= []).push(n.id);
  }

  const size = new Map<string, { w: number; lines: string[]; h: number }>();
  for (const n of nodes) {
    const lines = wrap(n.label);
    const w = Math.max(90, Math.min(280, Math.max(...lines.map((l) => l.length * 7 + 36))));
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
      laid.push({ id, label: nodes.find((n) => n.id === id)!.label, x: xCursor + (maxW - s.w) / 2, y, w: s.w, h: s.h, colX: xCursor, colW: maxW });
      y += s.h + Y_GAP;
    }
    xCursor += maxW + X_GAP;
    maxStackH = Math.max(maxStackH, y - Y_GAP);
  }

  return { nodes: laid, width: Math.max(xCursor - X_GAP + MARGIN, 10), height: Math.max(maxStackH + MARGIN, 10) };
}

export interface EdgeRoute {
  from: string;
  to: string;
  /** Knækpunkter (vinkelret rute: vandret → lodret → vandret). */
  points: { x: number; y: number }[];
  /** Returkant (del af en cyklus), som føres i en bane under diagrammet. */
  back: boolean;
}

/**
 * Vinkelrette kantruter. Alle segmenter er vandrette eller lodrette, så de kan tegnes
 * som simple linjer i PowerPoint (hvis API hverken kan vende linjer eller lave pilehoveder)
 * og ser ens ud i SVG-forhåndsvisningen og Word-billedet.
 */
export function routeEdges(d: DiagramResult, L: DiagramLayout = layoutDiagram(d)): EdgeRoute[] {
  const pos = new Map(L.nodes.map((n) => [n.id, n]));
  const routes: EdgeRoute[] = [];
  const bottom = Math.max(...L.nodes.map((n) => n.y + n.h));
  let lane = 0;
  for (const e of d.edges ?? []) {
    const a = pos.get(e.from);
    const b = pos.get(e.to);
    if (!a || !b || a === b) continue;
    const x1 = a.x + a.w;
    const y1 = a.y + a.h / 2;
    const y2 = b.y + b.h / 2;
    if (b.colX > a.colX) {
      // Fremad: knæk i mellemrummet lige foran målets søjle og ind i målets venstre side.
      const mx = b.colX - X_GAP / 2;
      routes.push({
        from: e.from,
        to: e.to,
        points: [
          { x: x1, y: y1 },
          { x: mx, y: y1 },
          { x: mx, y: y2 },
          { x: b.x, y: y2 },
        ],
        back: false,
      });
    } else {
      // Returkant (cyklus): ud til højre, ned i sin egen bane under diagrammet, tilbage og op
      // i mellemrummet foran målet, og ind i målets venstre side. Alle lodrette segmenter ligger
      // mellem søjlerne (en kvart afstand fra kanten – fremadkanter bruger midten), så de
      // krydser aldrig en knude.
      const ly = bottom + BACK_LANE * ++lane;
      const g = X_GAP / 4;
      const out = a.colX + a.colW + g;
      const into = b.colX - g;
      routes.push({
        from: e.from,
        to: e.to,
        points: [
          { x: x1, y: y1 },
          { x: out, y: y1 },
          { x: out, y: ly },
          { x: into, y: ly },
          { x: into, y: y2 },
          { x: b.x, y: y2 },
        ],
        back: true,
      });
    }
  }
  return routes;
}

/** Samlet udstrækning inkl. returkanternes baner (som kan ligge til venstre for og under knuderne). */
export function diagramBounds(L: DiagramLayout, routes: EdgeRoute[]): { x: number; y: number; width: number; height: number } {
  let minX = 0;
  let maxX = L.width;
  let maxY = L.height;
  for (const r of routes)
    for (const p of r.points) {
      minX = Math.min(minX, p.x - MARGIN);
      maxX = Math.max(maxX, p.x + MARGIN);
      maxY = Math.max(maxY, p.y + MARGIN);
    }
  return { x: minX, y: 0, width: maxX - minX, height: maxY };
}

const ARROW = 8;
/** Afstand mellem returkanternes baner under diagrammet. */
const BACK_LANE = 14;

/**
 * Renderer diagrammet som selvstændigt SVG (afrundede rektangler, tekst og pile).
 * Bruges til forhåndsvisning i task pane og som kilde til PNG (Word).
 */
export function diagramToSvg(d: DiagramResult): string {
  const L = layoutDiagram(d);
  const routes = routeEdges(d, L);
  const B = diagramBounds(L, routes);
  const parts: string[] = [];

  // Kanter først, så knuderne tegnes ovenpå.
  for (const r of routes) {
    parts.push(
      `<polyline points="${r.points.map((p) => `${p.x},${p.y}`).join(" ")}" fill="none" stroke="#6b7280" stroke-width="1.5" stroke-linejoin="round"/>`,
    );
    const end = r.points[r.points.length - 1];
    parts.push(`<polygon points="${end.x},${end.y} ${end.x - ARROW},${end.y - ARROW / 2} ${end.x - ARROW},${end.y + ARROW / 2}" fill="#6b7280"/>`);
  }

  for (const n of L.nodes) {
    const lines = wrap(n.label);
    const cx = n.x + n.w / 2;
    const y0 = n.y + n.h / 2 - ((lines.length - 1) * LINE_H) / 2 + 5;
    const tspans = lines.map((l, i) => `<tspan x="${cx}" ${i === 0 ? `y="${y0}"` : `dy="${LINE_H}"`}>${escapeXml(l)}</tspan>`).join("");
    parts.push(
      `<rect x="${n.x}" y="${n.y}" width="${n.w}" height="${n.h}" rx="8" fill="#eff6ff" stroke="#3b82f6"/>`,
      `<text text-anchor="middle" font-family="-apple-system, 'Segoe UI', sans-serif" font-size="13" fill="#111827">${tspans}</text>`,
    );
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${B.width}" height="${B.height}" viewBox="${B.x} ${B.y} ${B.width} ${B.height}">${parts.join("")}</svg>`;
}
