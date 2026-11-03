import type {
  AuditEntry,
  Endpoint,
  HealthInfo,
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

/**
 * Hvordan add-in'et når modellen (ADR-0007):
 *  - direct: add-in'et kalder selv endpointens OpenAI-API (standard for privat/SMB)
 *  - bridge: via den lokale MIAI Bridge (valgfri; offline, låst politik, audit-log)
 */
export type BackendMode = "direct" | "bridge";

/** Brugerdefineret opgave (F14) og dens oprindelse. */
export interface UserTaskInfo {
  spec: UserTaskFile;
  /** user = brugers mappe (redigerbar); machine = organisationen (read-only). */
  source: "user" | "machine";
}

export interface Backend {
  readonly mode: BackendMode;
  health(): Promise<HealthInfo>;
  models(endpointId?: string): Promise<ModelInfo[]>;
  settings(): Promise<SettingsView>;
  saveSettings(settings: Settings, policy?: Policy): Promise<void>;
  /** Det samlede opgavekatalog (indbygget + brugerdefineret); UI'et filtrerer på host. */
  tasks(): Promise<TaskView[]>;
  userTasks(): Promise<UserTaskInfo[]>;
  saveUserTask(spec: UserTaskFile): Promise<void>;
  deleteUserTask(id: string): Promise<void>;
  /** Tester en (evt. ikke gemt) endpoint: svarer den, og hvilke modeller har den? (#4) */
  probe(endpoint: Endpoint): Promise<ProbeResult>;
  /** Seneste audit-logposter (kun metadata). */
  audit(): Promise<AuditEntry[]>;
  /** Kører en opgave; kalder onEvent for hver hændelse. Kaster kun ved transportfejl. */
  runTask(id: string, req: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal): Promise<void>;
}
