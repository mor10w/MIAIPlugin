import type { HealthInfo, ModelInfo, Policy, Settings, SettingsView, TaskEvent, TaskRequest } from "@miai/shared";

// Task pane serveres af bridgen, så API'et ligger på samme origin.
const H = { "content-type": "application/json", "x-miai": "1" };

async function json<T>(res: Response): Promise<T> {
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error((body as { error?: string }).error ?? `${res.status} ${res.statusText}`);
  return body as T;
}

export const api = {
  health: () => fetch("/api/health", { headers: H }).then((r) => json<HealthInfo>(r)),
  models: (endpoint?: string) => fetch(`/api/models${endpoint ? `?endpoint=${encodeURIComponent(endpoint)}` : ""}`, { headers: H }).then((r) => json<ModelInfo[]>(r)),
  settings: () => fetch("/api/settings", { headers: H }).then((r) => json<SettingsView>(r)),
  saveSettings: (settings: Settings, policy?: Policy) =>
    fetch("/api/settings", { method: "PUT", headers: H, body: JSON.stringify({ settings, policy }) }).then((r) => json<{ ok: true }>(r)),
};

/** Kører en opgave og kalder onEvent for hver Server-Sent Event. */
export async function runTask(id: string, body: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal): Promise<void> {
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
