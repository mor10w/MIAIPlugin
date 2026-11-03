import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { homedir, platform } from "node:os";
import { dirname, join } from "node:path";
import { activeEndpoint, checkSettings, defaultSettings, isUrlAllowed, Policy, Settings } from "@miai/shared";

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
  /** Mappen config.json ligger i – audit.log og opgave-mappen (tasks/) ligger her. */
  readonly dir: string;
  private settings: Settings;
  private userPolicy: Policy | undefined;
  private machinePolicy: Policy | undefined;

  constructor(dir = configDir(), policyFile = machinePolicyFile()) {
    this.file = join(dir, "config.json");
    this.dir = dir;
    this.machinePolicy = readJson(policyFile, (raw) => Policy.parse(raw));
    if (!this.machinePolicy && existsSync(policyFile)) {
      // Fail closed: en ulæselig organisationspolitik må ikke åbne for flere hosts.
      this.machinePolicy = Policy.parse({ locked: true });
    }
    const stored = readJson(this.file, (raw) => {
      const r = raw as Partial<StoredConfig>;
      return { settings: Settings.parse(r.settings), policy: r.policy ? Policy.parse(r.policy) : undefined };
    });
    this.settings = stored?.settings ?? defaultSettings();
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
    const problem = checkSettings(next, effective, isUrlAllowed);
    if (problem) throw new PolicyError(problem);
    this.settings = next;
    if (policy) this.userPolicy = policy;
    this.save();
  }

  activeEndpoint() {
    return activeEndpoint(this.settings);
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
