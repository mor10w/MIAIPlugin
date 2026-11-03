/**
 * Kontekstbudget (#3): grov token-estimering, opdeling af lange tekster og beregning af,
 * hvor meget plads modellen har til svaret.
 *
 * Estimatet er bevidst konservativt (≈ 3,5 tegn pr. token for dansk/engelsk). Det bruges til
 * at planlægge – den faktiske afkortning opdages efterfølgende via modellens usage.
 */

export const CHARS_PER_TOKEN = 3.5;

/** Svarets mindste plads, når budgettet er stramt. */
export const MIN_OUTPUT_TOKENS = 256;

/** Del-størrelse ved opdeling, når modellens kontekst er ukendt (≈ 2.900 tokens – passer i 4k-kontekst). */
export const DEFAULT_CHUNK_CHARS = 10_000;

export function estimateTokens(text: string): number {
  return Math.ceil(text.length / CHARS_PER_TOKEN);
}

export function estimateMessages(messages: { content?: string | null }[]): number {
  // + lidt pr. besked til rolle og formattering
  return messages.reduce((n, m) => n + estimateTokens(m.content ?? "") + 4, 0);
}

/**
 * Hvor mange tokens svaret må bruge. Med kendt kontekst begrænses det, så prompt + svar passer
 * (vLLM m.fl. afviser ellers forespørgslen). Ukendt kontekst: brug indstillingen uændret.
 */
export function outputBudget(maxTokens: number, promptTokens: number, contextTokens?: number): number {
  if (!contextTokens) return maxTokens;
  return Math.max(MIN_OUTPUT_TOKENS, Math.min(maxTokens, contextTokens - promptTokens - 32));
}

/** Del-størrelse i tegn til map-reduce: plads til prompt-overhead og et kort del-resumé. */
export function chunkChars(contextTokens?: number): number {
  if (!contextTokens) return DEFAULT_CHUNK_CHARS;
  const room = contextTokens - 600 /* systemprompt m.m. */ - 700; /* del-resumé */
  return Math.max(2_000, Math.floor(room * CHARS_PER_TOKEN));
}

/**
 * Deler en tekst i dele på højst `maxChars` tegn – helst ved afsnit, ellers ved sætninger,
 * og kun som sidste udvej midt i en sætning.
 */
export function splitIntoChunks(text: string, maxChars: number): string[] {
  if (text.length <= maxChars) return [text];
  const chunks: string[] = [];
  let current = "";
  const flush = () => {
    if (current.trim()) chunks.push(current.trim());
    current = "";
  };
  const add = (piece: string, sep: string) => {
    if (current && current.length + sep.length + piece.length > maxChars) flush();
    current = current ? current + sep + piece : piece;
  };
  for (const para of text.split(/\n{2,}/)) {
    if (para.length <= maxChars) {
      add(para, "\n\n");
      continue;
    }
    // For langt afsnit: del ved sætninger, og hård-del sætninger, der selv er for lange.
    for (const sentence of para.split(/(?<=[.!?])\s+/)) {
      if (sentence.length <= maxChars) add(sentence, " ");
      else for (let i = 0; i < sentence.length; i += maxChars) add(sentence.slice(i, i + maxChars), "");
    }
  }
  flush();
  return chunks;
}

/**
 * Har modellen set markant mindre, end vi sendte? Sammenligner estimatet med modellens egen
 * usage.prompt_tokens. Værktøjer som Ollama afkorter prompts, der overstiger den konfigurerede
 * kontekst (num_ctx), uden at melde fejl.
 */
export function looksTruncated(estimatedPromptTokens: number, reportedPromptTokens: number | undefined): boolean {
  if (!reportedPromptTokens || estimatedPromptTokens < 1_500) return false;
  return reportedPromptTokens < estimatedPromptTokens * 0.6;
}
