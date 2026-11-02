import type { HealthInfo, ModelInfo, Policy, Settings, SettingsView, TaskEvent, TaskRequest } from "@miai/shared";
import type { Backend } from "./types";

// Task pane serveres af bridgen, så API'et ligger på samme origin.
const H = { "content-type": "application/json", "x-miai": "1" };

async function json<T>(res: Response): Promise<T> {
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error((body as { error?: string }).error ?? `${res.status} ${res.statusText}`);
  return body as T;
}

export class BridgeBackend implements Backend {
  readonly mode = "bridge" as const;

  health() {
    return fetch("/api/health", { headers: H }).then((r) => json<HealthInfo>(r));
  }
  models(endpoint?: string) {
    return fetch(`/api/models${endpoint ? `?endpoint=${encodeURIComponent(endpoint)}` : ""}`, { headers: H }).then((r) => json<ModelInfo[]>(r));
  }
  settings() {
    return fetch("/api/settings", { headers: H }).then((r) => json<SettingsView>(r));
  }
  async saveSettings(settings: Settings, policy?: Policy) {
    await fetch("/api/settings", { method: "PUT", headers: H, body: JSON.stringify({ settings, policy }) }).then((r) => json(r));
  }

  async runTask(id: string, body: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal): Promise<void> {
    const res = await fetch(`/api/tasks/${encodeURIComponent(id)}`, { method: "POST", headers: H, body: JSON.stringify(body), signal });
    if (!res.ok || !res.body) {
      await json(res);
      return;
    }
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = "";
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      let i: number;
      while ((i = buf.indexOf("\n\n")) >= 0) {
        const block = buf.slice(0, i);
        buf = buf.slice(i + 2);
        if (block.startsWith("data: ")) onEvent(JSON.parse(block.slice(6)) as TaskEvent);
      }
    }
  }
}

/** Er task pane serveret af en MIAI Bridge? */
export async function detectBridge(): Promise<boolean> {
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 1_500);
    const res = await fetch("/api/health", { headers: H, signal: ctrl.signal });
    clearTimeout(t);
    if (!res.ok) return false;
    const body = (await res.json()) as Partial<HealthInfo>;
    return typeof body.version === "string" && !!body.endpoint;
  } catch {
    return false;
  }
}
