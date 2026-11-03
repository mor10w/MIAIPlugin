/**
 * Detekterer degeneration i model-output: en kvantiseret Qwen3-tænkemodel kan falde i en
 * gentagelsessløjfe (fx "side 6 tabel side 6 tabel …"). Opdages tidligt, så strømmen kan
 * afbrydes i stedet for at køre til max_tokens.
 *
 * Signaturen er en *sammenhængende* gentagelse af slutningen. Ord, der blot optræder ofte
 * (fx "Ja" i en oversat tabelkolonne eller "færdig i team A, færdig i team B"), udløser
 * den ikke, fordi mønsteret skal gentage sig ord for ord helt ud til slutningen.
 */

/** Antal sammenhængende gentagelser, der kræves. En ægte sløjfe fortsætter, så det koster kun lidt ventetid. */
const MIN_REPEATS = 6;
/** Mindste samlede længde (ord) af den gentagne hale – beskytter mod korte, legitime gentagelser. */
const MIN_SPAN_WORDS = 12;
/** Mønstre på 1–6 ord. */
const MAX_PERIOD = 6;
/** Der kigges kun på slutningen af teksten, så kontrollen er O(1) pr. chunk. */
const TAIL_CHARS = 1_000;

export function isDegenerating(text: string): boolean {
  const words = text
    .slice(-TAIL_CHARS)
    .trim()
    .split(/\s+/)
    .map((w) => w.toLowerCase());
  for (let period = 1; period <= MAX_PERIOD; period++) {
    const need = Math.max(MIN_REPEATS, Math.ceil(MIN_SPAN_WORDS / period));
    if (words.length < period * need) break;
    let repeats = 1;
    // Sammenlign blok for blok bagfra med den sidste blok.
    for (let end = words.length - period; end - period >= 0; end -= period) {
      let same = true;
      for (let i = 0; i < period; i++) {
        if (words[end - period + i] !== words[words.length - period + i]) {
          same = false;
          break;
        }
      }
      if (!same) break;
      if (++repeats >= need) return true;
    }
  }
  return false;
}
