import { z } from "zod";
import { DocContext } from "@miai/shared";

/**
 * Eval-sættets format (#2). Hver case er én kørsel af en indbygget opgave med et fast input og
 * en række automatiske tjek. `rubric` bruges kun, hvis der køres med en dommermodel (--judge-model).
 */

export const Check = z.discriminatedUnion("type", [
  /** Alle (eller med any: mindst én) af værdierne skal stå i svaret. Uden forskel på store/små bogstaver. */
  z.object({ type: z.literal("contains"), values: z.array(z.string()).min(1), any: z.boolean().optional() }),
  z.object({ type: z.literal("notContains"), values: z.array(z.string()).min(1) }),
  z.object({ type: z.literal("regex"), pattern: z.string(), flags: z.string().optional(), negate: z.boolean().optional() }),
  /** Svarets sprog (stopord-heuristik – godt nok til at fange "svarede på engelsk"). */
  z.object({ type: z.literal("language"), lang: z.enum(["da", "en", "de"]) }),
  /** Længde i ord, evt. i forhold til inputtet (maxRatio = svarets ord / inputtets ord). */
  z.object({
    type: z.literal("length"),
    minWords: z.number().int().optional(),
    maxWords: z.number().int().optional(),
    maxRatio: z.number().optional(),
    minRatio: z.number().optional(),
  }),
  /** Antal punkttegn (linjer, der starter med -, * eller 1.). */
  z.object({ type: z.literal("bullets"), min: z.number().int().optional(), max: z.number().int().optional() }),
  /** Intet forord som "Her er…"/"Sure, here is…" og ingen kodeblokke rundt om svaret. */
  z.object({ type: z.literal("noPreamble") }),
  /** Strukturerede svar: værdien på `path` (fx "slides" eller "formula") og dens form. */
  z.object({
    type: z.literal("result"),
    path: z.string(),
    minItems: z.number().int().optional(),
    maxItems: z.number().int().optional(),
    pattern: z.string().optional(),
    flags: z.string().optional(),
  }),
  /** Grammatik: rettelser, der skal findes (erstatningen skal indeholde `fix`), og evt. maks. antal. */
  z.object({
    type: z.literal("edits"),
    expect: z.array(z.object({ fix: z.string() })).default([]),
    maxEdits: z.number().int().optional(),
  }),
  /** TSV (Excel-celler): antal kolonner pr. række og antal rækker. */
  z.object({ type: z.literal("tsv"), columns: z.number().int().optional(), rows: z.number().int().optional() }),
  /**
   * Agent-planer (#6): antal handlinger, hvilke slags der skal/ikke må være med, og om al
   * `find`/`anchor`-tekst faktisk findes i dokumentet (checkPlan).
   */
  z.object({
    type: z.literal("plan"),
    min: z.number().int().optional(),
    max: z.number().int().optional(),
    include: z.array(z.string()).optional(),
    exclude: z.array(z.string()).optional(),
    valid: z.boolean().optional(),
  }),
  /** Kaldte modellen web_search (true) – eller lod den være (false)? */
  z.object({ type: z.literal("tool"), called: z.boolean() }),
]);
export type Check = z.infer<typeof Check>;

export const EvalCase = z.object({
  id: z.string().regex(/^[a-z0-9-]+$/),
  /** Indbygget opgave, fx "summarize" eller "generate-formula". */
  task: z.string(),
  lang: z.enum(["da", "en"]),
  /** Fx "F1" – funktionsnummeret fra PLAN.md. */
  feature: z.string().optional(),
  note: z.string().optional(),
  context: DocContext,
  params: z.record(z.string(), z.string()).default({}),
  history: z.array(z.object({ role: z.enum(["user", "assistant"]), content: z.string() })).default([]),
  /** Giv modellen web_search-værktøjet (svarene er faste testdata, intet sendes på nettet). */
  tools: z.boolean().default(false),
  checks: z.array(Check).default([]),
  /** Bedømmelseskriterier til dommermodellen (1–5). */
  rubric: z.string().optional(),
});
export type EvalCase = z.infer<typeof EvalCase>;
export type EvalCaseInput = z.input<typeof EvalCase>;

export interface CheckResult {
  check: Check["type"];
  ok: boolean;
  detail?: string;
}

/** Resultatet af én case mod én model og én konfiguration. */
export interface CaseResult {
  caseId: string;
  task: string;
  lang: "da" | "en";
  model: string;
  /** Konfigurationens navn, fx "json_schema/pp0". */
  config: string;
  ok: boolean;
  error?: string;
  output: string;
  result?: unknown;
  checks: CheckResult[];
  /** Antal kald til /chat/completions. For strukturerede opgaver: 1 = gyldigt i første forsøg. */
  requests: number;
  structured: boolean;
  /** Strukturerede svar: gyldigt i første forsøg / efter genforsøg (undefined for fritekst). */
  firstTryValid?: boolean;
  validAfterRetry?: boolean;
  repetition: boolean;
  toolCalls: string[];
  warnings: string[];
  durationMs: number;
  completionTokens?: number;
  judge?: { score: number; reason: string };
}
