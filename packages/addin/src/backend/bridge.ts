import type {
  AuditEntry,
  Endpoint,
  ExportSheetRequest,
  HealthInfo,
  KnowledgeStatus,
  ProbeResult,
  ModelInfo,
  Policy,
  Settings,
  SettingsView,
  TaskEvent,
  TaskRequest,
  TaskView,
  UserTaskFile,
} from "@miai/shared";
import type { Backend, KnowledgeApi, UserTaskInfo } from "./types";

// Task pane serveres af bridgen, så API'et ligger på samme origin.
// Adgangstokenet (#11) står i task pane-adressen fra brugerens personlige manifest
// (…/taskpane.html?token=…); det sendes med hvert API-kald.
const token = (() => {
  try {
    return new URLSearchParams(window.location.search).get("token") ?? "";
  } catch {
    return "";
  }
})();
const H: Record<string, string> = { "content-type": "application/json", "x-miai": "1", ...(token ? { "x-miai-token": token } : {}) };

/** Fejl fra bridgen med dens kode (fx "TOKEN"), så UI'et kan forklare præcist, hvad der skal gøres. */
export class BridgeError extends Error {
  constructor(
    message: string,
    readonly code?: string,
  ) {
    super(message);
  }
}

async function json<T>(res: Response): Promise<T> {
  const body = (await res.json().catch(() => ({}))) as { error?: string; code?: string };
  if (!res.ok) throw new BridgeError(body.error ?? `${res.status} ${res.statusText}`, body.code);
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

  tasks() {
    return fetch("/api/tasks", { headers: H }).then((r) => json<TaskView[]>(r));
  }
  userTasks() {
    return fetch("/api/user-tasks", { headers: H }).then((r) => json<UserTaskInfo[]>(r));
  }
  async saveUserTask(spec: UserTaskFile) {
    await fetch(`/api/user-tasks/${encodeURIComponent(spec.id)}`, { method: "PUT", headers: H, body: JSON.stringify(spec) }).then((r) => json(r));
  }
  async deleteUserTask(id: string) {
    await fetch(`/api/user-tasks/${encodeURIComponent(id)}`, { method: "DELETE", headers: H }).then((r) => json(r));
  }
  probe(endpoint: Endpoint) {
    return fetch("/api/probe", { method: "POST", headers: H, body: JSON.stringify({ endpoint }) }).then((r) => json<ProbeResult>(r));
  }
  audit() {
    return fetch("/api/audit", { headers: H }).then((r) => json<AuditEntry[]>(r));
  }

  /** Eksport af et ark som .xlsx (tillæg ADR-0012). Bridgen kaster BridgeError, fx 409 ved eksisterende fil. */
  exportSheet(data: ExportSheetRequest) {
    return fetch("/api/export", { method: "POST", headers: H, body: JSON.stringify(data) }).then((r) => json<{ ok: true; path: string }>(r));
  }

  /** Tekst fra en vedhæftet PDF- eller Office-fil (#27), trukket ud af bridgen. */
  async extractFile(file: File): Promise<string> {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let bin = "";
    // I bidder: String.fromCharCode med hele filen på én gang rammer grænsen for argumenter.
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    const r = await fetch("/api/extract", { method: "POST", headers: H, body: JSON.stringify({ name: file.name, data: btoa(bin) }) });
    return (await json<{ text: string }>(r)).text;
  }

  readonly knowledge: KnowledgeApi = {
    status: () => fetch("/api/knowledge", { headers: H }).then((r) => json<KnowledgeStatus>(r)),
    reindex: () => fetch("/api/knowledge/reindex", { method: "POST", headers: H, body: "{}" }).then((r) => json<KnowledgeStatus>(r)),
    chunk: (id) => fetch(`/api/knowledge/chunks/${id}`, { headers: H }).then((r) => json<{ title: string; path: string; text: string }>(r)),
    open: async (chunkId) => {
      await fetch("/api/knowledge/open", { method: "POST", headers: H, body: JSON.stringify({ chunkId }) }).then((r) => json(r));
    },
  };

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
    // 401 med TOKEN = det er en bridge, men tokenet mangler (fx gammelt manifest). Brug bridgen,
    // så brugeren får den tydelige fejl i stedet for et forvirrende skift til direkte tilstand.
    if (res.status === 401) return ((await res.json().catch(() => ({}))) as { code?: string }).code === "TOKEN";
    if (!res.ok) return false;
    const body = (await res.json()) as Partial<HealthInfo>;
    return typeof body.version === "string" && !!body.endpoint;
  } catch {
    return false;
  }
}
