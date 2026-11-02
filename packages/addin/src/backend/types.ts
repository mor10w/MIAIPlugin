import type { HealthInfo, ModelInfo, Policy, Settings, SettingsView, TaskEvent, TaskRequest } from "@miai/shared";

/**
 * Hvordan add-in'et når modellen (ADR-0007):
 *  - direct: add-in'et kalder selv endpointens OpenAI-API (standard for privat/SMB)
 *  - bridge: via den lokale MIAI Bridge (valgfri; offline, låst politik, audit-log)
 */
export type BackendMode = "direct" | "bridge";

export interface Backend {
  readonly mode: BackendMode;
  health(): Promise<HealthInfo>;
  models(endpointId?: string): Promise<ModelInfo[]>;
  settings(): Promise<SettingsView>;
  saveSettings(settings: Settings, policy?: Policy): Promise<void>;
  /** Kører en opgave; kalder onEvent for hver hændelse. Kaster kun ved transportfejl. */
  runTask(id: string, req: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal): Promise<void>;
}
