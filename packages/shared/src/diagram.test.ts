import { describe, expect, it } from "vitest";
import { diagramToSvg, layoutDiagram, routeEdges, type DiagramResult } from "./diagram.js";

const sample: DiagramResult = {
  title: "Status",
  nodes: [
    { id: "n1", label: "Dataindsamling" },
    { id: "n2", label: "Analyse" },
    { id: "n3", label: "Rapport" },
  ],
  edges: [
    { from: "n1", to: "n2" },
    { from: "n2", to: "n3" },
  ],
};

describe("layoutDiagram", () => {
  it("placerer kæde i stigende x (venstre mod højre)", () => {
    const l = layoutDiagram(sample);
    const byId = new Map(l.nodes.map((n) => [n.id, n]));
    expect(byId.get("n1")!.x).toBeLessThan(byId.get("n2")!.x);
    expect(byId.get("n2")!.x).toBeLessThan(byId.get("n3")!.x);
  });

  it("er deterministisk", () => {
    expect(layoutDiagram(sample)).toEqual(layoutDiagram(sample));
  });

  it("har positive dimensioner og ingen negative koordinater", () => {
    const l = layoutDiagram(sample);
    expect(l.width).toBeGreaterThan(0);
    expect(l.height).toBeGreaterThan(0);
    for (const n of l.nodes) {
      expect(n.x).toBeGreaterThanOrEqual(0);
      expect(n.y).toBeGreaterThanOrEqual(0);
      expect(n.w).toBeGreaterThan(0);
      expect(n.h).toBeGreaterThan(0);
    }
  });

  it("håndterer cyklus uden at gå i uendelig sløjfe", () => {
    const cyclic: DiagramResult = {
      nodes: [
        { id: "a", label: "A" },
        { id: "b", label: "B" },
      ],
      edges: [
        { from: "a", to: "b" },
        { from: "b", to: "a" },
      ],
    };
    const l = layoutDiagram(cyclic);
    expect(l.nodes).toHaveLength(2);
  });

  it("en feedback-løkke til roden ødelægger ikke lagdelingen", () => {
    const d: DiagramResult = { ...sample, edges: [...sample.edges, { from: "n3", to: "n1" }] };
    const byId = new Map(layoutDiagram(d).nodes.map((n) => [n.id, n]));
    expect(byId.get("n1")!.x).toBeLessThan(byId.get("n2")!.x);
    expect(byId.get("n2")!.x).toBeLessThan(byId.get("n3")!.x);
    const routes = routeEdges(d);
    expect(routes.filter((r) => !r.back)).toHaveLength(2);
    expect(routes.filter((r) => r.back)).toHaveLength(1);
  });

  it("selvløkker ignoreres", () => {
    const d: DiagramResult = { nodes: [{ id: "a", label: "A" }], edges: [{ from: "a", to: "a" }] };
    expect(layoutDiagram(d).nodes).toHaveLength(1);
    expect(routeEdges(d)).toHaveLength(0);
  });

  it("ignorerer kanter med ukendte id'er", () => {
    const d: DiagramResult = { ...sample, edges: [...sample.edges, { from: "x", to: "y" }] };
    expect(() => layoutDiagram(d)).not.toThrow();
  });

  it("placerer to rødder i samme lag (vertikal stack)", () => {
    const d: DiagramResult = {
      nodes: [
        { id: "a", label: "A" },
        { id: "b", label: "B" },
        { id: "c", label: "C" },
      ],
      edges: [
        { from: "a", to: "c" },
        { from: "b", to: "c" },
      ],
    };
    const l = layoutDiagram(d);
    const byId = new Map(l.nodes.map((n) => [n.id, n]));
    expect(byId.get("a")!.x).toBeCloseTo(byId.get("b")!.x);
    expect(byId.get("a")!.y).not.toBeCloseTo(byId.get("b")!.y);
  });
});

describe("diagramToSvg", () => {
  it("indeholder alle labels og korrekte dimensioner", () => {
    const svg = diagramToSvg(sample);
    expect(svg).toContain("Dataindsamling");
    expect(svg).toContain("Analyse");
    expect(svg).toContain("Rapport");
    const l = layoutDiagram(sample);
    expect(svg).toContain(`width="${l.width}"`);
  });

  it("escaper XML-specialtegn", () => {
    const d: DiagramResult = { nodes: [{ id: "n1", label: 'A < "B" & C' }] };
    expect(diagramToSvg(d)).toContain("A &lt; &quot;B&quot; &amp; C");
  });

  it("tegner en kant og et pilehoved pr. edge", () => {
    const svg = diagramToSvg(sample);
    expect(svg.match(/<polyline /g)?.length).toBe(2);
    expect(svg.match(/<polygon /g)?.length).toBe(2);
  });

  it("ombryder lange labels som gyldige <tspan>-elementer", () => {
    const d: DiagramResult = { nodes: [{ id: "n1", label: "Indsamling af data fra alle afdelinger og eksterne leverandører" }] };
    const svg = diagramToSvg(d);
    expect(svg.match(/<tspan /g)?.length).toBeGreaterThan(1);
    expect(svg.match(/<tspan /g)?.length).toBe(svg.match(/<\/tspan>/g)?.length);
    expect(svg).not.toMatch(/[^<]tspan x=/); // ingen rå "tspan" som tekst
  });
});

describe("knudestørrelse", () => {
  it("bredden rummer den længste linje", () => {
    const d: DiagramResult = { nodes: [{ id: "n1", label: "Kort meget-meget-meget-lang-sammensætning" }] };
    const n = layoutDiagram(d).nodes[0];
    expect(n.w).toBeGreaterThanOrEqual(Math.min(280, "meget-meget-meget-lang-sammensætning".length * 7 + 36));
  });
});

describe("routeEdges", () => {
  it("laver vinkelrette ruter (kun vandrette/lodrette segmenter)", () => {
    for (const r of routeEdges(sample)) {
      for (let i = 1; i < r.points.length; i++) {
        const [p, q] = [r.points[i - 1], r.points[i]];
        expect(p.x === q.x || p.y === q.y).toBe(true);
      }
    }
  });

  it("alle kanter ender i målets venstre side; returkanter løber under diagrammet uden at krydse knuder", () => {
    const d: DiagramResult = {
      nodes: [
        { id: "a", label: "A" },
        { id: "b", label: "B" },
        { id: "c", label: "C" },
      ],
      edges: [
        { from: "a", to: "b" },
        { from: "b", to: "c" },
        { from: "c", to: "a" },
      ],
    };
    const L = layoutDiagram(d);
    const by = new Map(L.nodes.map((n) => [n.id, n]));
    for (const r of routeEdges(d, L)) {
      const end = r.points.at(-1)!;
      const t = by.get(r.to)!;
      expect(end.x).toBeCloseTo(t.x);
      if (r.back) expect(Math.max(...r.points.map((p) => p.y))).toBeGreaterThan(Math.max(...L.nodes.map((n) => n.y + n.h)));
      // Lodrette segmenter må ikke skære gennem en knude.
      for (let i = 1; i < r.points.length; i++) {
        const [p, q] = [r.points[i - 1], r.points[i]];
        if (p.x !== q.x) continue;
        for (const n of L.nodes) {
          const inside = p.x > n.x && p.x < n.x + n.w && Math.max(p.y, q.y) > n.y && Math.min(p.y, q.y) < n.y + n.h;
          expect(inside).toBe(false);
        }
      }
    }
  });
});
