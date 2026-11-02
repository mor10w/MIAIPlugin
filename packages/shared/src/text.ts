import type { GrammarResult } from "./tasks.js";

/**
 * Anvender grammatik-rettelser på en tekststreng (bruges hvor sporede ændringer ikke findes, fx PowerPoint).
 * Hver rettelse anvendes på første forekomst efter den forrige rettelse, så rækkefølgen bevares.
 * Returnerer ny tekst og antallet af rettelser, der ikke kunne placeres.
 */
export function applyEdits(text: string, edits: GrammarResult["edits"]): { text: string; skipped: number } {
  let out = "";
  let cursor = 0;
  let skipped = 0;
  for (const e of edits) {
    if (!e.original || e.original === e.replacement) continue;
    const idx = text.indexOf(e.original, cursor);
    if (idx < 0) {
      skipped++;
      continue;
    }
    out += text.slice(cursor, idx) + e.replacement;
    cursor = idx + e.original.length;
  }
  return { text: out + text.slice(cursor), skipped };
}

/**
 * Små modeller sætter ofte Markdown-fremhævning i svar, der skal være ren tekst.
 * Fjerner kun parvis **fed** og __fed__ – resten af teksten røres ikke.
 */
export function stripEmphasis(text: string): string {
  return text.replace(/\*\*(\S(?:[^*\n]*\S)?)\*\*/g, "$1").replace(/__(\S(?:[^_\n]*\S)?)__/g, "$1");
}

/** Parser TSV (som modellen returnerer til Excel) til et 2D-array. Fjerner afsluttende tom linje. */
export function parseTsv(tsv: string): string[][] {
  const lines = tsv.replace(/\r\n/g, "\n").replace(/\n+$/, "").split("\n");
  return lines.map((l) => l.split("\t"));
}

/** Gør et 2D-array rektangulært med den ønskede størrelse (fylder med "" / beskærer). */
export function fitGrid(grid: string[][], rows: number, cols: number): string[][] {
  return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => grid[r]?.[c] ?? ""));
}
