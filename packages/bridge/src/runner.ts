import { appendFile } from "node:fs/promises";
import { errorEvent, executeTask, type ParsedTaskRequest, type TaskDefinition, type TaskEvent } from "@miai/shared";
import { webSearch } from "./search.js";
import type { ConfigStore } from "./config.js";

interface RunOptions {
  task: TaskDefinition;
  req: ParsedTaskRequest;
  store: ConfigStore;
  emit: (e: TaskEvent) => void;
  signal: AbortSignal;
}

/** Værktøjskald fra modellen (ADR-0010). Kun web_search leveres – ukendte navne afvises. */
function makeToolExecutor(signal: AbortSignal) {
  return {
    specs: [
      {
        type: "function" as const,
        function: {
          name: "web_search",
          description: "Search the web for current information, news, prices, versions or facts you are not sure about. Returns a few results with title, URL and snippet.",
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
      return webSearch(query, signal);
    },
  };
}

/** Kører en opgave med den fælles motor og skriver evt. audit-log (kun metadata, aldrig indhold). */
export async function runTask({ task, req, store, emit, signal }: RunOptions): Promise<void> {
  const settings = store.getSettings();
  const ep = store.activeEndpoint();
  try {
    const summary = await executeTask(task, req, { settings, endpoint: ep, policy: store.getPolicy(), emit, signal, tools: makeToolExecutor(signal) });
    if (settings.auditLog) {
      const entry = {
        ts: new Date().toISOString(),
        task: task.id,
        endpoint: new URL(ep.baseUrl).host,
        model: summary.model,
        durationMs: summary.durationMs,
        inputChars: summary.inputChars,
        outputChars: summary.output.length,
        usage: summary.usage,
      };
      await appendFile(store.file.replace(/config\.json$/, "audit.log"), JSON.stringify(entry) + "\n").catch(() => {});
    }
  } catch (err) {
    if (!signal.aborted) emit(errorEvent(err));
  }
}
