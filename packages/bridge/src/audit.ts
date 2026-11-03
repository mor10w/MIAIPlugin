import { appendFile } from "node:fs/promises";
import { existsSync, readFileSync, renameSync } from "node:fs";
import { AuditEntry, MAX_AUDIT_FILE_BYTES } from "@miai/shared";

/**
 * Bridgens audit-log: JSONL (én AuditEntry pr. linje) i konfigurationsmappen.
 * Kun metadata – aldrig indhold (PLAN.md §7). Roteres ved MAX_AUDIT_FILE_BYTES:
 * audit.log → audit.log.old (maks. én generation).
 */

export function auditFile(configDir: string): string {
  return `${configDir}/audit.log`;
}

export async function appendAudit(file: string, entry: AuditEntry): Promise<void> {
  try {
    if (existsSync(file)) {
      const size = readFileSync(file).length;
      if (size > MAX_AUDIT_FILE_BYTES) {
        try {
          renameSync(file, `${file}.old`);
        } catch {
          // Fil i brug: behold den, og vedhæft fortsat.
        }
      }
    }
    await appendFile(file, JSON.stringify(entry) + "\n");
  } catch (err) {
    // Audit-log må aldrig bryde selve opgavekørslen.
    console.warn(`[miai] Kunne ikke skrive audit-log: ${(err as Error).message}`);
  }
}

/** De seneste poster (nederst i filen). Ugyldige linjer springes over. */
export async function readAuditTail(file: string, n = 100): Promise<AuditEntry[]> {
  if (!existsSync(file)) return [];
  let text: string;
  try {
    text = readFileSync(file, "utf8");
  } catch (err) {
    console.warn(`[miai] Kunne ikke læse audit-log: ${(err as Error).message}`);
    return [];
  }
  const lines = text.split("\n").filter((l) => l.trim());
  const out: AuditEntry[] = [];
  for (const line of lines.slice(-n)) {
    try {
      const e = AuditEntry.safeParse(JSON.parse(line));
      if (e.success) out.push(e.data);
    } catch {
      // Ugyldig linje: spring over.
    }
  }
  return out;
}
