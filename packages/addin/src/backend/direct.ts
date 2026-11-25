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
  MAX_LOCAL_AUDIT_ENTRIES,
  Endpoint,
  probeEndpoint,
  modelDetails,
  type ProbeResult,
  Policy,
  Settings,
  TASKS,
  TaskRequest,
  toTaskDefinition,
  toTaskView,
  UserTaskFile as UserTaskSchema,
  validateInput,
  AuditEntry as AuditSchema,
  type AuditEntry,
  type ClientOptions,
  type HealthInfo,
  type ModelDetails,
  type ModelInfo,
  type SettingsView,
  type TaskDefinition,
  type TaskEvent,
  type TaskView,
  type UserTaskFile,
  OrgConfig,
  ORG_CONFIG_FILE,
} from "@miai/shared";
import type { Backend, UserTaskInfo } from "./types";

const VERSION = "0.1.0";
const KEY_SETTINGS = "miai.settings.v1";
const KEY_POLICY = "miai.policy.v1";
/** Brugerdefinerede opgaver (F14): browseren kan ikke læse mapper, så filerne gemmes i localStorage. */
const KEY_USER_TASKS = "miai.userTasks.v1";
/** Audit-loggen som ringbuffer (kun metadata). */
const KEY_AUDIT = "miai.audit.v1";

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

/**
 * Henter organisationens konfiguration fra add-in'ets egen adresse (#7, ADR-0015). Ingen fil
 * (404) = ingen central konfiguration. Findes filen, men kan den ikke læses, lukkes der ned:
 * en låst politik med standardens tilladte hosts – en ødelagt fil må aldrig åbne for mere.
 */
export async function loadOrgConfig(base = window.location.href): Promise<OrgConfig | undefined> {
  let res: Response;
  try {
    res = await fetch(new URL(ORG_CONFIG_FILE, base), { cache: "no-store", signal: AbortSignal.timeout(3_000) });
  } catch {
    return undefined; // offline eller blokeret: som uden fil
  }
  if (res.status === 404) return undefined;
  try {
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return OrgConfig.parse(await res.json());
  } catch (err) {
    console.warn(`[miai] ${ORG_CONFIG_FILE} kunne ikke læses – politikken låses: ${(err as Error).message}`);
    return { policy: Policy.parse({ locked: true }) };
  }
}

/** Add-in'et taler selv med endpointens OpenAI-API. Ingen lokal service nødvendig. */
export class DirectBackend implements Backend {
  readonly mode = "direct" as const;
  private client: ClientOptions = { unreachableHint: corsHint() };
  /** Organisationens politik (låst) og standardindstillinger, hvis add-in'et er udrullet centralt. */
  private orgPolicy: Policy | undefined;
  private orgSettings: Settings | undefined;

  constructor(org?: OrgConfig) {
    if (org?.policy) this.orgPolicy = Policy.parse({ ...org.policy, locked: true });
    if (org?.settings) {
      try {
        this.orgSettings = Settings.parse({ ...defaultSettings(), ...org.settings });
      } catch (err) {
        console.warn(`[miai] Ugyldige standardindstillinger i ${ORG_CONFIG_FILE}: ${(err as Error).message}`);
      }
    }
  }

  /** Opretter backenden med organisationens konfiguration, hvis den findes. */
  static async create(): Promise<DirectBackend> {
    return new DirectBackend(await loadOrgConfig());
  }

  private get settingsNow(): Settings {
    return load(
      KEY_SETTINGS,
      (r) => Settings.parse(r),
      () => this.orgSettings ?? defaultSettings(),
    );
  }
  private get policyNow(): Policy {
    return this.orgPolicy ?? load(KEY_POLICY, (r) => Policy.parse(r), defaultPolicy);
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
    const host = epHost(ep.baseUrl);
    const started = Date.now();
    try {
      const models = await cachedModels(ep, timeout(5_000), this.client);
      this.pushAudit({ ts: new Date().toISOString(), kind: "models", host, status: "ok", durationMs: Date.now() - started });
      return models.map((id) => ({ id, endpoint: ep.id }));
    } catch (err) {
      this.pushAudit({ ts: new Date().toISOString(), kind: "models", host, status: "error" });
      throw err;
    }
  }

  async settings(): Promise<SettingsView> {
    const hasPolicy = !!store.get(KEY_POLICY);
    return { settings: this.settingsNow, policy: this.policyNow, policySource: this.orgPolicy ? "organization" : hasPolicy ? "browser" : "default" };
  }

  async saveSettings(settings: Settings, policy?: Policy): Promise<void> {
    const s = Settings.parse(settings);
    if (policy && (this.orgPolicy || this.policyNow.locked)) throw new Error("Politikken er låst af din organisation og kan ikke ændres.");
    const p = policy ? Policy.parse(policy) : this.policyNow;
    const problem = checkSettings(s, p, isUrlAllowed);
    if (problem) throw new Error(problem);
    store.set(KEY_SETTINGS, JSON.stringify(s));
    if (policy) store.set(KEY_POLICY, JSON.stringify(p));
    clearModelCache();
  }

  async tasks(): Promise<TaskView[]> {
    const views = TASKS.map((t) => toTaskView(t));
    for (const c of this.loadCustom()) views.push(toTaskView(toTaskDefinition(c.spec), true));
    return views;
  }

  async userTasks(): Promise<UserTaskInfo[]> {
    return this.loadCustom();
  }

  async saveUserTask(spec: UserTaskFile): Promise<void> {
    const parsed = UserTaskSchema.parse(spec);
    // I direkte tilstand er der ingen maskin-mappe, så kollision med en indbygget id
    // kan ikke overstyres – brugeren skal vælge et andet id. (Bridgen tillader det.)
    if (getTask(parsed.id)) throw new Error(`Id'en '${parsed.id}' er reserveret af en indbygget opgave.`);
    const all = this.loadCustom().map((c) => c.spec);
    const i = all.findIndex((x) => x.id === parsed.id);
    if (i >= 0) all[i] = parsed;
    else all.push(parsed);
    store.set(KEY_USER_TASKS, JSON.stringify(all));
  }

  async deleteUserTask(id: string): Promise<void> {
    const all = this.loadCustom().map((c) => c.spec);
    const rest = all.filter((x) => x.id !== id);
    if (rest.length === all.length) throw new Error("Opgaven findes ikke.");
    store.set(KEY_USER_TASKS, JSON.stringify(rest));
  }

  async modelDetails(endpoint: Endpoint, model: string): Promise<ModelDetails> {
    const ep = Endpoint.parse(endpoint);
    if (!isUrlAllowed(ep.baseUrl, this.policyNow)) throw new Error("Endpoint er ikke tilladt af politikken.");
    // Fra Indstillinger: spørg altid værktøjet igen (fx efter at konteksten er ændret).
    return modelDetails(ep, model, { signal: timeout(5_000), client: this.client, fresh: true });
  }

  async probe(endpoint: Endpoint): Promise<ProbeResult> {
    return probeEndpoint(Endpoint.parse(endpoint), this.policyNow, {
      signal: timeout(6_000),
      client: this.client,
      // Et no-cors-kald lykkes (opakt svar), når noget lytter – så mangler kun CORS-opsætningen.
      reachableWithoutCors: async (url) => {
        await fetch(url, { mode: "no-cors", signal: timeout(3_000) });
        return true;
      },
    });
  }

  async audit(): Promise<AuditEntry[]> {
    return this.loadAudit();
  }

  /** Indbygget opgave, evt. suppleret med brugerdefinerede fra localStorage. */
  private resolveTask(id: string): TaskDefinition | undefined {
    const builtin = getTask(id);
    if (builtin) return builtin;
    const spec = this.loadCustom().find((c) => c.spec.id === id)?.spec;
    return spec ? toTaskDefinition(spec) : undefined;
  }

  private loadCustom(): UserTaskInfo[] {
    const out: UserTaskInfo[] = [];
    for (const s of load(
      KEY_USER_TASKS,
      (r) => (Array.isArray(r) ? r : []),
      () => [],
    )) {
      const p = UserTaskSchema.safeParse(s);
      if (p.success) out.push({ spec: p.data, source: "user" });
    }
    return out;
  }

  private loadAudit(): AuditEntry[] {
    const out: AuditEntry[] = [];
    for (const e of load(
      KEY_AUDIT,
      (r) => (Array.isArray(r) ? r : []),
      () => [],
    )) {
      const p = AuditSchema.safeParse(e);
      if (p.success) out.push(p.data);
    }
    return out;
  }

  /** Append til ringbufferen. Skriver intet, hvis audit-loggen er slået fra. */
  private pushAudit(entry: AuditEntry): void {
    if (!this.settingsNow.auditLog) return;
    const all = [...this.loadAudit(), entry].slice(-MAX_LOCAL_AUDIT_ENTRIES);
    store.set(KEY_AUDIT, JSON.stringify(all));
  }

  async runTask(id: string, body: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal): Promise<void> {
    const task = this.resolveTask(id);
    if (!task) throw new Error(`Ukendt opgave '${id}'`);
    const req = TaskRequest.parse(body);
    validateInput(task, req);
    const settings = this.settingsNow;
    const ep = activeEndpoint(settings);
    try {
      const summary = await executeTask(task, req, { settings, endpoint: ep, policy: this.policyNow, emit: onEvent, signal, client: this.client });
      if (settings.auditLog) {
        this.pushAudit({
          ts: new Date().toISOString(),
          kind: "task",
          task: id,
          host: epHost(ep.baseUrl),
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
        onEvent(e);
        if (settings.auditLog && e.type === "error")
          this.pushAudit({ ts: new Date().toISOString(), kind: "task", task: id, host: epHost(ep.baseUrl), status: "error", code: e.code });
      } else throw err;
    }
  }
}

/** Host inkl. port ud fra basis-URL'en – uden at bryde kørslen ved ulæselig URL. */
function epHost(baseUrl: string): string {
  try {
    return new URL(baseUrl).host;
  } catch {
    return baseUrl;
  }
}

function timeout(ms: number): AbortSignal {
  const c = new AbortController();
  setTimeout(() => c.abort(), ms);
  return c.signal;
}
