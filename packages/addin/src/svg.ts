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

/** Gærdede blokke i tekst-rækkefølge (lang = tagget efter ```; open = intet afsluttende gærde). */
interface Fenced {
  lang: string;
  code: string;
  open: boolean;
}

function fencedBlocks(text: string): Fenced[] {
  const blocks: Fenced[] = [];
  const re = /```([A-Za-z0-9_+-]*)[ \t]*\r?\n?/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const start = m.index + m[0].length;
    const close = text.indexOf("```", start);
    if (close === -1) {
      blocks.push({ lang: m[1], code: text.slice(start), open: true });
      break;
    }
    blocks.push({ lang: m[1], code: text.slice(start, close), open: false });
    re.lastIndex = close + 3;
  }
  return blocks;
}

/**
 * Første komplette <svg>…</svg> i teksten. En SVG uden </svg> (afbrudt svar) kan ikke
 * rasteriseres og tæller ikke – ellers ville knappen "Indsæt som billede" altid fejle.
 */
function completeSvg(text: string): string | null {
  return /<svg[\s>][\s\S]*?<\/svg>/i.exec(text)?.[0] ?? null;
}

/**
 * Finder den første ```mermaid- eller ```svg-blok i teksten (i tekst-rækkefølge).
 * Fallbacks for små modeller, der gærder sløset: en SVG under et andet tag (fx ```xml eller ```html,
 * også i en ulukket blok), en ulukket ```mermaid-blok (afbrudt svar) og rå <svg>…</svg> uden gærde.
 * SVG renses altid af kalderen (sanitizeSvg) før brug.
 */
export function extractGraphic(text: string): GraphicBlock | null {
  const mermaid = /```mermaid[ \t]*\r?\n([\s\S]*?)```/i.exec(text);
  const svg = /```svg[ \t]*\r?\n([\s\S]*?)```/i.exec(text);
  if (mermaid && svg) {
    return mermaid.index <= svg.index ? { kind: "mermaid", code: mermaid[1].trim() } : { kind: "svg", code: svg[1].trim() };
  }
  if (mermaid) return { kind: "mermaid", code: mermaid[1].trim() };
  if (svg) return { kind: "svg", code: svg[1].trim() };

  const blocks = fencedBlocks(text);
  for (const b of blocks) {
    const code = completeSvg(b.code);
    if (code) return { kind: "svg", code };
  }
  // Ulukket ```mermaid-blok: koden er måske ufuldstændig, men mermaid kan ofte tegne det, der er.
  for (const b of blocks) {
    if (b.open && /^mermaid$/i.test(b.lang) && b.code.trim()) return { kind: "mermaid", code: b.code.trim() };
  }
  const bare = completeSvg(text);
  return bare ? { kind: "svg", code: bare } : null;
}
