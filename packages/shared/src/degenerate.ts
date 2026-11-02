/**
 * Detekterer degeneration i model-output: en kvantiseret Qwen3-tænkemodel kan ved lav
 * temperatur falde i en gentagelsessløjfe (fx "side 6 / tabel / side 6 …"). Her opdages
 * det tidligt, så strømmen kan afbrydes i stedet for at køre til max_tokens.
 */

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Returnerer true, hvis de sidste ord i `text` udgør et kort mønster (1–3 ord),
 * der gentages mindst 5 gange i de sidste 80 ord.
 */
export function isDegenerating(text: string): boolean {
  const words = text.trim().split(/\s+/);
  if (words.length < 16) return false;
  const window = words.slice(-80).join(" ");
  for (const len of [1, 2, 3]) {
    const pat = words.slice(-len).join(" ");
    if (pat.length < 2 || pat.length > 30) continue;
    const re = new RegExp(`\\b${escapeRe(pat)}\\b`, "g");
    const count = (window.match(re) ?? []).length;
    if (count >= 5) return true;
  }
  return false;
}
