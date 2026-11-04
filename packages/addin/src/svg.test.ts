import { describe, expect, it } from "vitest";
import { extractGraphic, extractSvgBlock, sanitizeSvg } from "./svg.js";

describe("extractSvgBlock", () => {
  it("finder ```svg-blok uden gærde", () => {
    const text = 'Her er figuren:\n```svg\n<svg viewBox="0 0 100 50"><rect/></svg>\n```\nhåb';
    expect(extractSvgBlock(text)).toBe('<svg viewBox="0 0 100 50"><rect/></svg>');
  });

  it("returnerer null uden blok", () => {
    expect(extractSvgBlock("ingen figur")).toBeNull();
  });

  it("returnerer null, hvis blokken er ulukket", () => {
    expect(extractSvgBlock("```svg\n<svg")).toBeNull();
  });
});

describe("sanitizeSvg", () => {
  it("fjerner <script>-elementer", () => {
    const svg = "<svg><rect/><script>alert(1)</script></svg>";
    expect(sanitizeSvg(svg)).not.toContain("alert");
  });

  it("fjerner event-håndlere (dbl- og enkel citation)", () => {
    const svg = "<svg><rect onclick=\"evil()\"/><circle onmouseover='x()'/></svg>";
    const clean = sanitizeSvg(svg);
    expect(clean).not.toContain("onclick");
    expect(clean).not.toContain("onmouseover");
  });

  it("fjerner javascript:- og data:-URL'er i href", () => {
    const svg = '<svg><a href="javascript:evil()"/><use xlink:href="data:image/svg+xml;base64,xxx"/></svg>';
    const clean = sanitizeSvg(svg);
    expect(clean).not.toContain("javascript:");
    expect(clean).not.toContain("data:image/svg+xml;base64,xxx");
  });

  it("fjerner <foreignObject> (indlejret HTML)", () => {
    const svg = '<svg><rect/><foreignObject><iframe src="https://evil.example"></iframe></foreignObject></svg>';
    const clean = sanitizeSvg(svg);
    expect(clean).not.toContain("foreignObject");
    expect(clean).not.toContain("iframe");
  });

  it("bevarer ufarlig SVG", () => {
    const svg = '<svg viewBox="0 0 10 10"><rect width="5" height="5" fill="red"/></svg>';
    expect(sanitizeSvg(svg)).toBe(svg);
  });
});

describe("extractGraphic", () => {
  const mermaid = "```mermaid\ngraph TD; A-->B;\n```";
  const svg = "```svg\n<svg></svg>\n```";

  it("foretrækker den tidligste blok (mermaid først)", () => {
    expect(extractGraphic(`før\n${mermaid}\nmellem\n${svg}`)).toEqual({ kind: "mermaid", code: "graph TD; A-->B;" });
  });

  it("foretrækker den tidligste blok (svg først)", () => {
    expect(extractGraphic(`før\n${svg}\nmellem\n${mermaid}`)).toEqual({ kind: "svg", code: "<svg></svg>" });
  });

  it("returnerer den type, der findes", () => {
    expect(extractGraphic(mermaid)).toEqual({ kind: "mermaid", code: "graph TD; A-->B;" });
    expect(extractGraphic(svg)).toEqual({ kind: "svg", code: "<svg></svg>" });
  });

  it("returnerer null uden blok", () => {
    expect(extractGraphic("kun tekst")).toBeNull();
  });

  it("finder SVG under et andet gærde-tag (fx ```xml)", () => {
    const text = 'Her er figuren:\n```xml\n<svg viewBox="0 0 10 5"><rect/></svg>\n```';
    expect(extractGraphic(text)).toEqual({ kind: "svg", code: '<svg viewBox="0 0 10 5"><rect/></svg>' });
  });

  it("finder en komplet SVG i en ulukket ```svg-blok (afbrudt efter figuren)", () => {
    const text = 'Som SVG:\n```svg\n<svg viewBox="0 0 10 5"><rect/></svg>\nForklaring, der blev afbrudt';
    expect(extractGraphic(text)).toEqual({ kind: "svg", code: '<svg viewBox="0 0 10 5"><rect/></svg>' });
  });

  it("ignorerer en afbrudt SVG uden </svg> (kan ikke rasteriseres)", () => {
    expect(extractGraphic('Som SVG:\n```svg\n<svg viewBox="0 0 10 5"><rect/>')).toBeNull();
    expect(extractGraphic('<svg viewBox="0 0 10 5"><rect/>')).toBeNull();
  });

  it("tager kun <svg>…</svg> fra en ```xml- eller ```html-blok med omgivende markup", () => {
    const text = '```html\n<div><svg viewBox="0 0 10 5"><rect/></svg></div>\n```';
    expect(extractGraphic(text)).toEqual({ kind: "svg", code: '<svg viewBox="0 0 10 5"><rect/></svg>' });
  });

  it("tager kun den første af flere rå SVG'er", () => {
    const text = '<svg viewBox="0 0 1 1"><rect/></svg> og <svg viewBox="0 0 2 2"><circle/></svg>';
    expect(extractGraphic(text)).toEqual({ kind: "svg", code: '<svg viewBox="0 0 1 1"><rect/></svg>' });
  });

  it("finder mermaid i en ulukket ```mermaid-blok (afbrudt svar)", () => {
    const text = "```mermaid\ngraph TD; A-->B";
    expect(extractGraphic(text)).toEqual({ kind: "mermaid", code: "graph TD; A-->B" });
  });

  it("finder rå <svg>-markup uden gærde", () => {
    const text = 'Her er din struktur:\n<svg viewBox="0 0 100 50"><rect/></svg>\nhåb';
    expect(extractGraphic(text)).toEqual({ kind: "svg", code: '<svg viewBox="0 0 100 50"><rect/></svg>' });
  });

  it("foretrækker lukkede blokke for fallbacks", () => {
    const text = `rå <svg></svg> markup først\n${mermaid}`;
    expect(extractGraphic(text)).toEqual({ kind: "mermaid", code: "graph TD; A-->B;" });
  });

  it("ignorerer rå markup, der ikke er en komplet SVG", () => {
    expect(extractGraphic("jeg kan hjælpe med <svg> og <rect>, hvis du vil")).toBeNull();
  });

  it("ignorerer ```-blokke uden SVG", () => {
    expect(extractGraphic("```text\nbare tekst\n```")).toBeNull();
  });
});
