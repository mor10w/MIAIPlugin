import { z } from "zod";

/**
 * Audit-logpost (PLAN.md: "audit-log over alle udgående kald – host, model, tokens, ikke indhold").
 * Kun metadata: aldrig dokumenttekst, prompts eller svarkontent.
 */
export const AuditEntry = z.object({
  /** ISO-8601-tidsstempel. */
  ts: z.string(),
  /**
   * task = en opgavekørsel; models = henting af modeller fra /v1/models; search = web-søgning (ADR-0010);
   * index = indeksering af mapper (F15); retrieve = opslag i vidensindekset (F15).
   */
  kind: z.enum(["task", "models", "search", "index", "retrieve"]),
  /** Opgave-id (fx "grammar"), kun ved kind=task. */
  task: z.string().max(80).optional(),
  /** Endpointens host inkl. port, fx "127.0.0.1:11434". */
  host: z.string().max(200),
  /** Modelnavn, hvis kendt. */
  model: z.string().max(200).optional(),
  status: z.enum(["ok", "error"]),
  /** Fejlkode (fx "POLICY", "NO_MODEL"), kun ved status=error. */
  code: z.string().max(40).optional(),
  durationMs: z.number().int().nonnegative().max(3_600_000).optional(),
  inputChars: z.number().int().nonnegative().max(1_000_000).optional(),
  outputChars: z.number().int().nonnegative().max(1_000_000).optional(),
  promptTokens: z.number().int().nonnegative().max(1_000_000).optional(),
  completionTokens: z.number().int().nonnegative().max(1_000_000).optional(),
  /** index: antal filer indekseret/fjernet; retrieve: antal uddrag fundet. Aldrig stier eller indhold. */
  count: z.number().int().nonnegative().max(10_000_000).optional(),
});
export type AuditEntry = z.infer<typeof AuditEntry>;

/** Maks. poster i den lokale (browser) audit-log – ringbuffer i localStorage. */
export const MAX_LOCAL_AUDIT_ENTRIES = 200;

/** Maks. størrelse på bridgens audit.log, før den roteres (audit.log → audit.log.old). */
export const MAX_AUDIT_FILE_BYTES = 1_048_576;
