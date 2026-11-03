/**
 * Heuristik: lyder en chat-forespørgsel om at *ændre* dokumentet?
 * Hvis ja, foreslår chatten agent-tilstanden, så redigeringen sker i
 * dokumentet i stedet for at blive liggende som chat-tekst.
 *
 * JavaScripts \b kender kun a–z, så "Ændr" eller "Opdatér" ville aldrig matche. Ordgrænser er
 * derfor skrevet med Unicode-bogstaver (\p{L}) og flaget `u`.
 */

/** Ordstart og ordslut, der også virker for æ, ø, å og accenter. */
const W0 = "(?<![\\p{L}\\p{N}])";
const W1 = "(?![\\p{L}\\p{N}])";

// Dansk: bydeform og infinitiv af tydelige redigeringsverber (ikke "ret", der også betyder "temmelig").
const DA_VERBS = [
  "tilføj(?:e|er)?",
  "indsæt(?:te|ter)?",
  "ændr(?:e|er)?",
  "slet(?:te|ter)?",
  "fjern(?:e|er)?",
  "opdat[eé]r(?:e)?",
  "udvid(?:e|er)?",
  "forkort(?:e|er)?",
  "omskriv(?:e|er)?",
  "omformuler(?:e)?",
  "supplér|supplere?",
  "erstat(?:te|ter)?",
  "flyt(?:te|ter)?",
  "korrigér|korrigere?",
  "rediger(?:e)?",
];
// Danske vendinger: "gør teksten kortere", "skriv … ind i dokumentet", "lav afsnit 2 om til …", "ret … til …".
const DA_PHRASES = [
  "gør\\s+[\\p{L}\\s]{0,30}?(?:kortere|længere|enklere|tydeligere|mere\\s+\\p{L}+|mindre\\s+\\p{L}+|til\\s+(?:en\\s+)?(?:overskrift|punkt|liste|tabel))",
  "skriv\\s+[^?]{0,60}?(?:i|ind\\s+i|til)\\s+(?:dokumentet|teksten|rapporten|afsnittet)",
  "lav\\s+[^?]{0,40}?\\s+om(?:\\s+til)?",
  "ret\\s+[^?]{0,60}?\\s+til\\s+",
  // "… mangler en konklusion – skriv den"
  "skriv\\s+(?:den|det|dem)(?![\\p{L}])",
];
const EN = [
  "add(?:s|ed|ing)?",
  "insert(?:s|ed|ing)?",
  "edit(?:s|ed|ing)?",
  "remov(?:e|es|ed|ing)",
  "delet(?:e|es|ed|ing)",
  "rewrit(?:e|es|ing)|rewritten",
  "rephras(?:e|es|ed|ing)",
  "shorten(?:s|ed|ing)?",
  "replac(?:e|es|ed|ing)",
  "make\\s+[\\p{L}\\s]{0,30}?(?:shorter|longer|simpler|clearer|a\\s+(?:heading|bullet|list|table))",
];

const DA_RE = new RegExp(`${W0}(?:${DA_VERBS.join("|")})${W1}|${W0}(?:${DA_PHRASES.join("|")})`, "iu");
// "change"/"update" kun som verbum foran et objekt ("change the …", "update this …"), ikke "climate change" eller "latest update".
const EN_RE = new RegExp(`${W0}(?:${EN.join("|")})${W1}|${W0}(?:change|update|expand)\\s+(?:the|this|that|it|my|our|all|every)${W1}`, "iu");

/**
 * @returns true, hvis teksten indeholder et verbum eller en vending, der tydeligt peger på en dokumentændring.
 */
export function looksLikeEditRequest(prompt: string): boolean {
  return DA_RE.test(prompt) || EN_RE.test(prompt);
}
