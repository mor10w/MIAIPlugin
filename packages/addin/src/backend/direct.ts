import {
  activeEndpoint,
  cachedModels,
  checkSettings,
  clearModelCache,
  defaultPolicy,
  defaultSettings,
  errorEvent,
  executeTask,
  getTask,
  isUrlAllowed,
  Policy,
  Settings,
  TaskRequest,
  validateInput,
  type ClientOptions,
  type HealthInfo,
  type ModelInfo,
  type SettingsView,
  type TaskEvent,
} from "@miai/shared";
import type { Backend } from "./types";

const VERSION = "0.1.0";
const KEY_SETTINGS = "miai.settings.v1";
const KEY_POLICY = "miai.policy.v1";

/**
 * localStorage pr. add-in-origin. Kan være utilgængelig eller blive ryddet (fx privat tilstand,
 * visse Mac-builds), så alle kald er beskyttet, og der falder tilbage til hukommelsen.
 */
const memory = new Map<string, string>();
const store = {
  get(key: string): string | null {
    try {
      return window.localStorage.getItem(key) ?? memory.get(key) ?? null;
    } catch {
      return memory.get(key) ?? null;
    }
  },
  set(key: string, value: string) {
    memory.set(key, value);
    try {
      window.localStorage.setItem(key, value);
    } catch {
      /* kun i hukommelsen */
    }
  },
};

function load<T>(key: string, parse: (raw: unknown) => T, fallback: () => T): T {
  const raw = store.get(key);
  if (!raw) return fallback();
  try {
    return parse(JSON.parse(raw));
  } catch {
    return fallback();
  }
}

/** Forklaring ved forbindelsesfejl: i direkte tilstand er CORS den typiske årsag. */
export function corsHint(origin = window.location.origin): string {
  return (
    `Tjek at værktøjet tillader CORS fra ${origin}. ` +
    `Ollama: sæt OLLAMA_ORIGINS=${origin} (https://localhost:* er tilladt som standard). ` +
    `LM Studio: slå "Enable CORS" til under Developer/Server. ` +
    `vLLM: --allowed-origins. Alternativt: brug MIAI Bridge.`
  );
}

/** Add-in'et taler selv med endpointens OpenAI-API. Ingen lokal service nødvendig. */
export class DirectBackend implements Backend {
  readonly mode = "direct" as const;
  private client: ClientOptions = { unreachableHint: corsHint() };

  private get settingsNow(): Settings {
    return load(KEY_SETTINGS, (r) => Settings.parse(r), defaultSettings);
  }
  private get policyNow(): Policy {
    return load(KEY_POLICY, (r) => Policy.parse(r), defaultPolicy);
  }

  async health(): Promise<HealthInfo> {
    const ep = activeEndpoint(this.settingsNow);
    try {
      if (!isUrlAllowed(ep.baseUrl, this.policyNow)) throw new Error("Endpoint er ikke tilladt af politikken.");
      await cachedModels(ep, timeout(3_000), this.client);
      return { ok: true, version: VERSION, endpoint: { id: ep.id, baseUrl: ep.baseUrl, reachable: true } };
    } catch (err) {
      return { ok: true, version: VERSION, endpoint: { id: ep.id, baseUrl: ep.baseUrl, reachable: false, error: (err as Error).message } };
    }
  }

  async models(endpointId?: string): Promise<ModelInfo[]> {
    const s = this.settingsNow;
    const ep = s.endpoints.find((e) => e.id === (endpointId ?? s.activeEndpoint)) ?? activeEndpoint(s);
    if (!isUrlAllowed(ep.baseUrl, this.policyNow)) throw new Error("Endpoint er ikke tilladt af politikken.");
    const models = await cachedModels(ep, timeout(5_000), this.client);
    return models.map((id) => ({ id, endpoint: ep.id }));
  }

  async settings(): Promise<SettingsView> {
    const hasPolicy = !!store.get(KEY_POLICY);
    return { settings: this.settingsNow, policy: this.policyNow, policySource: hasPolicy ? "browser" : "default" };
  }

  async saveSettings(settings: Settings, policy?: Policy): Promise<void> {
    const s = Settings.parse(settings);
    const p = policy ? Policy.parse(policy) : this.policyNow;
    const problem = checkSettings(s, p, isUrlAllowed);
    if (problem) throw new Error(problem);
    store.set(KEY_SETTINGS, JSON.stringify(s));
    if (policy) store.set(KEY_POLICY, JSON.stringify(p));
    clearModelCache();
  }

  async runTask(id: string, body: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal): Promise<void> {
    const task = getTask(id);
    if (!task) throw new Error(`Ukendt opgave '${id}'`);
    const req = TaskRequest.parse(body);
    validateInput(task, req);
    const settings = this.settingsNow;
    try {
      await executeTask(task, req, { settings, endpoint: activeEndpoint(settings), policy: this.policyNow, emit: onEvent, signal, client: this.client });
    } catch (err) {
      if (!signal.aborted) onEvent(errorEvent(err));
      else throw err;
    }
  }
}

function timeout(ms: number): AbortSignal {
  const c = new AbortController();
  setTimeout(() => c.abort(), ms);
  return c.signal;
}
