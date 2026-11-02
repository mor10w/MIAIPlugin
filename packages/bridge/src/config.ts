import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { homedir, platform } from "node:os";
import { dirname, join } from "node:path";
import { Policy, Settings } from "@miai/shared";
import { isUrlAllowed } from "./policy.js";

export const DEFAULT_SETTINGS: Settings = Settings.parse({
  endpoints: [
    {
      id: "ollama",
      name: "Ollama",
      baseUrl: "http://127.0.0.1:11434/v1",
      structuredOutput: "json_schema",
    },
  ],
  activeEndpoint: "ollama",
});

export function configDir(): string {
  if (process.env.MIAI_CONFIG_DIR) return process.env.MIAI_CONFIG_DIR;
  switch (platform()) {
    case "win32":
      return join(process.env.APPDATA ?? join(homedir(), "AppData", "Roaming"), "MIAI");
    case "darwin":
      return join(homedir(), "Library", "Application Support", "MIAI");
    default:
      return join(process.env.XDG_CONFIG_HOME ?? join(homedir(), ".config"), "miai");
  }
}

/** Maskin-politik sat af en administrator. Har forrang og kan ikke ændres via UI. */
export function machinePolicyFile(): string {
  if (process.env.MIAI_POLICY_FILE) return process.env.MIAI_POLICY_FILE;
  switch (platform()) {
    case "win32":
      return join(process.env.ProgramData ?? "C:\\ProgramData", "MIAI", "policy.json");
    case "darwin":
      return "/Library/Application Support/MIAI/policy.json";
    default:
      return "/etc/miai/policy.json";
  }
}

interface StoredConfig {
  settings: Settings;
  policy?: Policy;
}

export class ConfigStore {
  readonly file: string;
  private settings: Settings;
  private userPolicy: Policy | undefined;
  private machinePolicy: Policy | undefined;

  constructor(dir = configDir(), policyFile = machinePolicyFile()) {
    this.file = join(dir, "config.json");
    this.machinePolicy = readJson(policyFile, (raw) => Policy.parse(raw));
    if (!this.machinePolicy && existsSync(policyFile)) {
      // Fail closed: en ulæselig organisationspolitik må ikke åbne for flere hosts.
      this.machinePolicy = Policy.parse({ locked: true });
    }
    const stored = readJson(this.file, (raw) => {
      const r = raw as Partial<StoredConfig>;
      return { settings: Settings.parse(r.settings), policy: r.policy ? Policy.parse(r.policy) : undefined };
    });
    this.settings = stored?.settings ?? structuredClone(DEFAULT_SETTINGS);
    this.userPolicy = stored?.policy;
  }

  get policySource(): "default" | "user" | "machine" {
    return this.machinePolicy ? "machine" : this.userPolicy ? "user" : "default";
  }

  getPolicy(): Policy {
    return this.machinePolicy ?? this.userPolicy ?? Policy.parse({});
  }

  getSettings(): Settings {
    return this.settings;
  }

  /** Validerer og gemmer. Kaster en fejl, hvis en endpoint ikke er tilladt af politikken. */
  update(next: Settings, policy?: Policy): void {
    let effective = this.getPolicy();
    if (policy) {
      if (this.machinePolicy || effective.locked) throw new PolicyError("Politikken er låst af din organisation og kan ikke ændres.");
      effective = policy;
    }
    if (!next.endpoints.some((e) => e.id === next.activeEndpoint)) throw new PolicyError("Aktiv endpoint findes ikke.");
    const ids = new Set<string>();
    for (const e of next.endpoints) {
      if (ids.has(e.id)) throw new PolicyError(`Endpoint-id '${e.id}' er brugt to gange.`);
      ids.add(e.id);
      if (!isUrlAllowed(e.baseUrl, effective)) {
        throw new PolicyError(`Host for '${e.name}' (${e.baseUrl}) er ikke tilladt af politikken. Tilføj hosten under "Tilladte hosts".`);
      }
    }
    this.settings = next;
    if (policy) this.userPolicy = policy;
    this.save();
  }

  activeEndpoint() {
    const s = this.settings;
    return s.endpoints.find((e) => e.id === s.activeEndpoint) ?? s.endpoints[0];
  }

  private save(): void {
    mkdirSync(dirname(this.file), { recursive: true });
    const tmp = `${this.file}.tmp`;
    const data: StoredConfig = { settings: this.settings, policy: this.userPolicy };
    writeFileSync(tmp, JSON.stringify(data, null, 2), { mode: 0o600 });
    renameSync(tmp, this.file);
  }
}

export class PolicyError extends Error {}

function readJson<T>(file: string, parse: (raw: unknown) => T): T | undefined {
  if (!existsSync(file)) return undefined;
  try {
    return parse(JSON.parse(readFileSync(file, "utf8")));
  } catch (err) {
    console.warn(`[miai] Ignorerer ugyldig fil ${file}: ${(err as Error).message}`);
    return undefined;
  }
}
