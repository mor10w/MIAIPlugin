/**
 * SVG-støtte (ADR-0009): modellen kan tegne direkte i SVG. Blokken udtrækkes,
 * renses (ingen scripts, event-håndlere eller eksterne referencer – output kan
 * være påvirket af prompt injection i dokumenttekst) og rasteriseres til PNG.
 */

/** Returnerer indholdet af den første ```svg-blok i teksten (uden gærde), ellers null. */
export function extractSvgBlock(text: string): string | null {
  const m = /```svg[ \t]*\r?\n([\s\S]*?)```/i.exec(text);
  return m ? m[1].trim() : null;
}

/**
 * Rens SVG til lokal visning/rasterisering: fjerner scripts, event-håndlere,
 * javascript:/data:-URL'er og indlejret HTML (foreignObject).
 */
export function sanitizeSvg(svg: string): string {
  return svg
    .replace(/<script\b[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+\s*=\s*"[^"]*"/gi, "")
    .replace(/\son\w+\s*=\s*'[^']*'/gi, "")
    .replace(/(href|xlink:href)\s*=\s*(["'])\s*(?:javascript|data):[^"']*\2/gi, "")
    .replace(/<foreignObject\b[\s\S]*?<\/foreignObject>/gi, "");
}

export type GraphicBlock = { kind: "mermaid" | "svg"; code: string };

/** Finder den første ```mermaid- eller ```svg-blok i teksten (i tekst-rækkefølge). */
export function extractGraphic(text: string): GraphicBlock | null {
  const mermaid = /```mermaid[ \t]*\r?\n([\s\S]*?)```/i.exec(text);
  const svg = /```svg[ \t]*\r?\n([\s\S]*?)```/i.exec(text);
  if (mermaid && svg) {
    return mermaid.index <= svg.index ? { kind: "mermaid", code: mermaid[1].trim() } : { kind: "svg", code: svg[1].trim() };
  }
  if (mermaid) return { kind: "mermaid", code: mermaid[1].trim() };
  if (svg) return { kind: "svg", code: svg[1].trim() };
  return null;
}
