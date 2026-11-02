import { describe, expect, it } from "vitest";
import { diagramToSvg, layoutDiagram, type DiagramResult } from "./diagram.js";

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

  it("tegner en kant pr. edge", () => {
    const svg = diagramToSvg(sample);
    expect(svg.match(/d="M [\d.]+ [\d.]+ C /g)?.length).toBe(2); // to buede kanter (pilehoveder medregnes ikke)
  });
});
