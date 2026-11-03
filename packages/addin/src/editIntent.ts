/**
 * Heuristik: lyder en chat-forespørgsel om at *ændre* dokumentet?
 * Hvis ja, foreslår chatten agent-tilstanden, så redigeringen sker i
 * dokumentet i stedet for at blive liggende som chat-tekst.
 */

// Dansk: ord med fleksion (tilføj/tilføje, indsæt/indsætte …). "ret" kun som hele ord.
const DA = /\b(tilføj\w*|tilføje\w*|indsæt\w*|ret\b|rette\b|ændr\w*|slet\w*|opdater\w*|udvid\w*|forkort\w*|omskriv\w*|suppler\w*|mangl\w*)\b/i;
// Engelsk: stammer dækker flektion (added, editing, removing …).
const EN = /\b(add\w*|insert\w*|edit\w*|change\w*|update\w*|remov\w*|delet\w*|rewrit\w*|expand\w*|shorten\w*|missing)\b/i;

/**
 * @returns true, hvis teksten indeholder et verbum/ord, der tydeligt peger på en dokumentændring.
 */
export function looksLikeEditRequest(prompt: string): boolean {
  return DA.test(prompt) || EN.test(prompt);
}
