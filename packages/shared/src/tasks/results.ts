/** JSON-skemaer for opgavernes strukturerede svar (#60). */
import { z } from "zod";

export const GrammarResult = z.object({
  edits: z
    .array(
      z.object({
        original: z.string().describe("Exact text from the document that should change, short (a few words), copied verbatim"),
        replacement: z.string().describe("The corrected text"),
        explanation: z.string().describe("Very short reason, in the document's language"),
      }),
    )
    .describe("Empty if the text has no errors"),
});
export type GrammarResult = z.infer<typeof GrammarResult>;

export const SlidesResult = z.object({
  slides: z
    .array(
      z.object({
        title: z.string(),
        bullets: z.array(z.string()).describe("3-6 short bullet points"),
        notes: z.string().describe("Speaker notes, 1-3 sentences").optional(),
      }),
    )
    .min(1),
});
export type SlidesResult = z.infer<typeof SlidesResult>;

export const FormulaResult = z.object({
  formula: z.string().describe("Excel formula starting with '='. English function names, comma as argument separator."),
  explanation: z.string().describe("Short explanation in the user's language"),
});
export type FormulaResult = z.infer<typeof FormulaResult>;
