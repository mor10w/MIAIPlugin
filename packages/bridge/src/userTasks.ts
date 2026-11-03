import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { TASKS, toTaskDefinition, UserTaskFile, type Settings, type TaskDefinition } from "@miai/shared";
import { configDir } from "./config.js";

/**
 * Brugerdefinerede opgaver (F14): JSON-filer fra mapper, så opgaver deles ved at dele mappen.
 * To lag: en maskin-mappe (MIAI_TASKS_DIR, sat af administratoren) og en bruger-mappe
 * (indstillingen customTasksDir, standard <config>/tasks). Maskin-filer vinder ved id-kollision.
 */

export type TaskSource = "user" | "machine";

export interface LoadedTask {
  spec: UserTaskFile;
  source: TaskSource;
  /** Filen opgaven blev læst fra (bruges ved sletning – filnavnet behøver ikke være <id>.json). */
  file: string;
}

/** Maskin-mappe sat af administratoren. Udefineret = ingen organisation-opgaver. */
export function machineTasksDir(): string | undefined {
  const d = process.env.MIAI_TASKS_DIR;
  return d ? resolve(d) : undefined;
}

/** Bruger-mappe: indstillingen customTasksDir, ellers <config>/tasks. */
export function userTasksDir(settings: Settings, base = configDir()): string {
  return settings.customTasksDir ? resolve(settings.customTasksDir) : join(base, "tasks");
}

function jsonFiles(dir: string): string[] {
  if (!existsSync(dir)) return [];
  const out: string[] = [];
  const walk = (d: string): void => {
    let entries;
    try {
      entries = readdirSync(d, { withFileTypes: true });
    } catch (err) {
      console.warn(`[miai] Kan ikke læse opgave-mappe ${d}: ${(err as Error).message}`);
      return;
    }
    for (const entry of entries) {
      const p = join(d, entry.name);
      if (entry.isDirectory()) walk(p);
      else if (entry.isFile() && entry.name.endsWith(".json")) out.push(p);
    }
  };
  walk(dir);
  return out;
}

function loadDir(dir: string, source: TaskSource): LoadedTask[] {
  const out: LoadedTask[] = [];
  for (const file of jsonFiles(dir)) {
    try {
      const r = UserTaskFile.safeParse(JSON.parse(readFileSync(file, "utf8")));
      if (r.success) out.push({ spec: r.data, source, file });
      else console.warn(`[miai] Ugyldig opgavefil ${file}: ${r.error.issues.map((i) => i.message).join(", ")}`);
    } catch (err) {
      console.warn(`[miai] Kan ikke læse opgavefil ${file}: ${(err as Error).message}`);
    }
  }
  return out;
}

// Lille TTL-cache, så en delt mappe (fx OneDrive) opdateres uden disk-læsning ved hvert kald.
const CACHE_MS = 5_000;
let cache: { key: string; at: number; value: Map<string, LoadedTask> } | undefined;

/** Alle brugerdefinerede opgaver fra begge mapper (maskin vinder ved id-kollision). */
export function loadUserTasks(settings: Settings, base = configDir()): Map<string, LoadedTask> {
  const key = `${userTasksDir(settings, base)}|${machineTasksDir() ?? ""}`;
  if (cache && cache.key === key && Date.now() - cache.at < CACHE_MS) return cache.value;
  const map = new Map<string, LoadedTask>();
  for (const t of loadDir(userTasksDir(settings, base), "user")) map.set(t.spec.id, t);
  const machine = machineTasksDir();
  if (machine) for (const t of loadDir(machine, "machine")) map.set(t.spec.id, t);
  cache = { key, at: Date.now(), value: map };
  return map;
}

/** Kaldes efter skrivehandlinger, så ændringer ses med det samme. */
export function invalidateUserTasks(): void {
  cache = undefined;
}

/** Indbyggede + brugerdefinerede. En brugeropgave med samme id som en indbygget overskriver den (fx organisationsprompts). */
export function mergedCatalog(settings: Settings, base = configDir()): { defs: TaskDefinition[]; customIds: Set<string> } {
  const user = loadUserTasks(settings, base);
  const all = new Map<string, TaskDefinition>(TASKS.map((t) => [t.id, t]));
  for (const { spec } of user.values()) all.set(spec.id, toTaskDefinition(spec));
  return { defs: [...all.values()], customIds: new Set(user.keys()) };
}
