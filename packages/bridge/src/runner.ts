import { appendAudit, auditFile } from "./audit.js";
import {
  CELL_ASK_TASK,
  effectiveSearch,
  errorEvent,
  executeTask,
  isUrlAllowed,
  type AuditEntry,
  type Policy,
  type SearchConfig,
  type ParsedTaskRequest,
  type TaskDefinition,
  type TaskEvent,
} from "@miai/shared";
import { searchWeb } from "./search.js";
import type { ConfigStore } from "./config.js";
import type { KnowledgeIndex } from "./knowledge/index.js";

interface RunOptions {
  task: TaskDefinition;
  req: ParsedTaskRequest;
  store: ConfigStore;
  emit: (e: TaskEvent) => void;
  signal: AbortSignal;
  /** Vidensindekset (F15). Bruges af chat, når brugeren har slået "Brug mine mapper" til. */
  knowledge?: KnowledgeIndex;
}

/**
 * Spørgsmålet, der søges med: prompten – og ved korte opfølgninger ("og i Sverige?") også
 * brugerens forrige besked, så søgningen har nok at gå efter.
 */
function retrievalQuery(req: ParsedTaskRequest, taskId?: string): string {
  const prompt = (req.params.prompt ?? "").trim();
  // Regnearksfunktionen: prompten er ofte generel ("Hvad er kundens rabat?"), og cellens input siger hvem.
  if (taskId === CELL_ASK_TASK) return `${prompt}\n${(req.context.table || req.context.text).slice(0, 500)}`.trim();
  if (prompt.length >= 40) return prompt;
  const prev = [...req.history].reverse().find((m) => m.role === "user")?.content ?? "";
  return `${prev}\n${prompt}`.trim();
}

/** Bygger en audit-post (kun metadata) ud fra modelkaldets resultat. */
function baseEntry(task: TaskDefinition, epHost: string): Pick<AuditEntry, "ts" | "kind" | "task" | "host"> {
  return { ts: new Date().toISOString(), kind: "task", task: task.id, host: epHost };
}

/** Værktøjskald fra modellen (ADR-0010). Kun web_search leveres – ukendte navne afvises. */
function makeToolExecutor(
  signal: AbortSignal,
  search: SearchConfig,
  policy: Policy,
  onSearch: (ok: boolean, durationMs: number, host: string) => Promise<void>,
) {
  return {
    specs: [
      {
        type: "function" as const,
        function: {
          name: "web_search",
          description:
            "Search the web for current information, news, prices, versions or facts you are not sure about. Returns a few results with title, URL and snippet.",
          parameters: {
            type: "object",
            properties: { query: { type: "string", description: "The search query, in the user's language or English" } },
            required: ["query"],
          },
        },
      },
    ],
    async run(name: string, argsJson: string): Promise<string> {
      if (name !== "web_search") return `Ukendt værktøj: ${name}`;
      let query = "";
      try {
        query = ((JSON.parse(argsJson) as { query?: string }).query ?? "").trim();
      } catch {
        // Ugyldige argumenter fra modellen.
      }
      if (!query) return "Manglende søgeforespørgsel i argumenterne.";
      const started = Date.now();
      const r = await searchWeb(search, query, (url) => isUrlAllowed(url, policy), signal);
      await onSearch(r.ok, Date.now() - started, r.host);
      return r.text;
    },
  };
}

/** Kører en opgave med den fælles motor og skriver evt. audit-log (kun metadata, aldrig indhold). */
export async function runTask({ task, req: clientReq, store, emit, signal, knowledge }: RunOptions): Promise<void> {
  const settings = store.getSettings();
  const ep = store.activeEndpoint();
  let host: string;
  try {
    host = new URL(ep.baseUrl).host;
  } catch {
    host = ep.baseUrl; // skal ikke kunne bryde kørslen, hvis URL'en er ulæselig
  }
  try {
    // Søgninger er udgående kald til en ekstern tjeneste og logges derfor også (kun host og tid).
    const onSearch = async (ok: boolean, durationMs: number, searchHost: string) => {
      if (settings.auditLog)
        await appendAudit(auditFile(store.dir), {
          ts: new Date().toISOString(),
          kind: "search",
          task: task.id,
          host: searchHost,
          status: ok ? "ok" : "error",
          durationMs,
        });
    };
    // Kilder kommer kun fra bridgens eget indeks – aldrig fra klienten.
    const req: ParsedTaskRequest = { ...clientReq, sources: [] };
    const policy = store.getPolicy();
    // Chatten og regnearksfunktionen uden optimized (#80) kan slå op i brugerens egne mapper.
    const usesKnowledge = task.id === "chat" || (task.id === CELL_ASK_TASK && clientReq.params.optimized !== "on");
    if (usesKnowledge && knowledge && clientReq.params.useKnowledge === "true" && settings.knowledge.enabled && policy.knowledge) {
      const started = Date.now();
      try {
        const found = await knowledge.search(retrievalQuery(clientReq, task.id), undefined, signal);
        req.sources = found.sources;
        emit({ type: "sources", sources: found.refs });
        if (settings.auditLog)
          await appendAudit(auditFile(store.dir), {
            ts: new Date().toISOString(),
            kind: "retrieve",
            task: task.id,
            host,
            model: settings.knowledge.embeddingModel,
            status: "ok",
            durationMs: Date.now() - started,
            count: found.sources.length,
          });
      } catch (err) {
        if (settings.auditLog)
          await appendAudit(auditFile(store.dir), { ts: new Date().toISOString(), kind: "retrieve", task: task.id, host, status: "error" });
        throw err;
      }
    }
    const summary = await executeTask(task, req, {
      settings,
      endpoint: ep,
      policy: store.getPolicy(),
      emit,
      signal,
      tools: makeToolExecutor(signal, effectiveSearch(settings, store.getPolicy()), store.getPolicy(), onSearch),
    });
    if (settings.auditLog) {
      await appendAudit(auditFile(store.dir), {
        ...baseEntry(task, host),
        model: summary.model,
        status: "ok",
        durationMs: summary.durationMs,
        inputChars: summary.inputChars,
        outputChars: summary.output.length,
        promptTokens: summary.usage?.promptTokens,
        completionTokens: summary.usage?.completionTokens,
      });
    }
  } catch (err) {
    if (!signal.aborted) {
      const e = errorEvent(err);
      emit(e);
      if (settings.auditLog && e.type === "error") {
        await appendAudit(auditFile(store.dir), { ...baseEntry(task, host), status: "error", code: e.code });
      }
    }
  }
}
