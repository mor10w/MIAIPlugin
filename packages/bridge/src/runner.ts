import { appendFile } from "node:fs/promises";
import { errorEvent, executeTask, type ParsedTaskRequest, type TaskDefinition, type TaskEvent } from "@miai/shared";
import type { ConfigStore } from "./config.js";

interface RunOptions {
  task: TaskDefinition;
  req: ParsedTaskRequest;
  store: ConfigStore;
  emit: (e: TaskEvent) => void;
  signal: AbortSignal;
}

/** Kører en opgave med den fælles motor og skriver evt. audit-log (kun metadata, aldrig indhold). */
export async function runTask({ task, req, store, emit, signal }: RunOptions): Promise<void> {
  const settings = store.getSettings();
  const ep = store.activeEndpoint();
  try {
    const summary = await executeTask(task, req, { settings, endpoint: ep, policy: store.getPolicy(), emit, signal });
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
