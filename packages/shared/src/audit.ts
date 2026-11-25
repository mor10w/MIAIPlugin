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
   * index = indeksering af mapper (F15); retrieve = opslag i vidensindekset (F15);
   * export = et ark gemt som .xlsx-fil (tillæg ADR-0012; host = mappen, fx "desktop").
   */
  kind: z.enum(["task", "models", "search", "index", "retrieve", "export"]),
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
  /**
   * index: antal filer indekseret/fjernet; retrieve: antal uddrag fundet; task = cell-ask: antal
   * celler samlet i posten (#80). Aldrig stier eller indhold.
   */
  count: z.number().int().nonnegative().max(10_000_000).optional(),
});
export type AuditEntry = z.infer<typeof AuditEntry>;

/** Maks. poster i den lokale (browser) audit-log – ringbuffer i localStorage. */
export const MAX_LOCAL_AUDIT_ENTRIES = 200;

/** Maks. størrelse på bridgens audit.log, før den roteres (audit.log → audit.log.old). */
export const MAX_AUDIT_FILE_BYTES = 1_048_576;

/** Cellekald, der ligger tættere end dette, hører til samme kørsel. */
const CELL_RUN_WINDOW_MS = 60_000;

/**
 * Regnearksfunktionen =MIAI.ASK() (#80) kan give tusind kald i én kørsel, og den lokale log rummer
 * kun MAX_LOCAL_AUDIT_ENTRIES poster. Hører `entry` til samme kørsel som den seneste post (samme
 * opgave, host, model og udfald, kort efter hinanden), returneres én samlet post; ellers undefined.
 */
export function mergeCellAudit(last: AuditEntry | undefined, entry: AuditEntry): AuditEntry | undefined {
  if (!last || entry.kind !== "task" || entry.task !== "cell-ask") return undefined;
  if (last.kind !== "task" || last.task !== entry.task || last.host !== entry.host || last.model !== entry.model) return undefined;
  if (last.status !== entry.status || last.code !== entry.code) return undefined;
  const gap = Date.parse(entry.ts) - Date.parse(last.ts);
  if (!(gap >= 0 && gap <= CELL_RUN_WINDOW_MS)) return undefined;
  const sum = (a: number | undefined, b: number | undefined, max: number) =>
    a === undefined && b === undefined ? undefined : Math.min(max, (a ?? 0) + (b ?? 0));
  return {
    ...entry,
    count: (last.count ?? 1) + (entry.count ?? 1),
    durationMs: sum(last.durationMs, entry.durationMs, 3_600_000),
    inputChars: sum(last.inputChars, entry.inputChars, 1_000_000),
    outputChars: sum(last.outputChars, entry.outputChars, 1_000_000),
    promptTokens: sum(last.promptTokens, entry.promptTokens, 1_000_000),
    completionTokens: sum(last.completionTokens, entry.completionTokens, 1_000_000),
  };
}
