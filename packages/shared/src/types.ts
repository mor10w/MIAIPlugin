import { z } from "zod";

export const HostKind = z.enum(["word", "excel", "powerpoint", "unknown"]);
export type HostKind = z.infer<typeof HostKind>;

/** Det udsnit af dokumentet, som en opgave arbejder på. Bygges af add-in'ets host-adapter. */
export const DocContext = z.object({
  host: HostKind,
  /** Markeret tekst (eller hele dokumentet, hvis intet er markeret og opgaven tillader det). */
  text: z.string().max(200_000).default(""),
  /** true hvis `text` er hele dokumentet frem for en markering. */
  isWholeDocument: z.boolean().default(false),
  /** Excel: adressen på markeringen, fx "Ark1!C5". */
  address: z.string().optional(),
  /** Excel: formlen i markeringens øverste venstre celle. */
  formula: z.string().optional(),
  /** Excel: markeringen som TSV (værdier), begrænset i størrelse. */
  table: z.string().max(100_000).optional(),
  /** Excel: kolonneoverskrifter eller andet omkringliggende indhold, der hjælper modellen. */
  surroundings: z.string().max(20_000).optional(),
});
export type DocContext = z.infer<typeof DocContext>;

export const ChatMessage = z.object({
  role: z.enum(["system", "user", "assistant"]),
  content: z.string(),
});
export type ChatMessage = z.infer<typeof ChatMessage>;

export const TaskRequest = z.object({
  context: DocContext,
  /** Opgaveparametre, fx { targetLanguage: "engelsk" } eller { tone: "formel" }. */
  params: z.record(z.string(), z.string().max(4_000)).default({}),
  /** Kun for chat: tidligere beskeder i samtalen. */
  history: z.array(ChatMessage).max(100).default([]),
  /** Overstyr model for denne forespørgsel. */
  model: z.string().optional(),
});
export type TaskRequest = z.input<typeof TaskRequest>;
export type ParsedTaskRequest = z.output<typeof TaskRequest>;

/** Hændelser bridgen sender som Server-Sent Events. */
export type TaskEvent =
  | { type: "meta"; model: string; endpoint: string }
  | { type: "delta"; text: string }
  /** Fremdrift for strukturerede (JSON) svar, som ikke vises løbende. */
  | { type: "progress"; chars: number }
  | { type: "result"; data: unknown }
  | { type: "error"; message: string; code?: string }
  | {
      type: "done";
      usage?: { promptTokens?: number; completionTokens?: number };
      /** Sat, hvis motoren afbrød svaret, fx fordi modellen gik i en gentagelsessløjfe. */
      stopReason?: "repetition";
    };

// ---------------------------------------------------------------------------
// Indstillinger og politik
// ---------------------------------------------------------------------------

export const StructuredOutputMode = z.enum(["json_schema", "json_object", "none"]);
export type StructuredOutputMode = z.infer<typeof StructuredOutputMode>;

export const Endpoint = z.object({
  id: z.string().regex(/^[a-z0-9-]{1,40}$/),
  name: z.string().min(1).max(80),
  /** Basis-URL inkl. /v1, fx http://127.0.0.1:11434/v1 */
  baseUrl: z.string().url(),
  apiKey: z.string().max(500).default(""),
  structuredOutput: StructuredOutputMode.default("json_schema"),
  stripThinking: z.boolean().default(true),
  /** Ekstra felter, der flettes ind i hver chat/completions-forespørgsel (fx { "num_ctx": 8192 }). */
  extraBody: z.record(z.string(), z.unknown()).default({}),
  timeoutMs: z.number().int().min(1_000).max(600_000).default(120_000),
});
export type Endpoint = z.infer<typeof Endpoint>;

export const Settings = z.object({
  endpoints: z.array(Endpoint).min(1),
  activeEndpoint: z.string(),
  /** Standardmodel. Tom = første model, endpointen rapporterer. */
  defaultModel: z.string().default(""),
  /** Model pr. opgave-id, fx { "explain-formula": "qwen3:14b" }. */
  taskModels: z.record(z.string(), z.string()).default({}),
  temperature: z.number().min(0).max(2).default(0.3),
  maxTokens: z.number().int().min(16).max(65_536).default(8_192),
  uiLanguage: z.enum(["da", "en"]).default("da"),
  auditLog: z.boolean().default(false),
});
export type Settings = z.infer<typeof Settings>;

export const Policy = z.object({
  /** Hostnavne/IP'er bridgen må kalde. Understøtter "*.domæne" wildcard. */
  allowedHosts: z.array(z.string()).default(["127.0.0.1", "localhost", "::1"]),
  /** Låst politik kan ikke ændres via UI (sat af organisationen). */
  locked: z.boolean().default(false),
});
export type Policy = z.infer<typeof Policy>;

/** Det, add-in'et får fra GET /api/settings (API-nøgler maskeres). */
export interface SettingsView {
  settings: Settings;
  policy: Policy;
  /** browser = gemt i add-in'et (direkte tilstand); machine = låst af organisationen (kun bridge). */
  policySource: "default" | "user" | "machine" | "browser";
}

export interface ModelInfo {
  id: string;
  endpoint: string;
}

export interface HealthInfo {
  ok: boolean;
  version: string;
  endpoint: { id: string; baseUrl: string; reachable: boolean; error?: string };
}
