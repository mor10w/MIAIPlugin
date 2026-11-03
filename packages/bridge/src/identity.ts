import { randomBytes, timingSafeEqual } from "node:crypto";
import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { createServer } from "node:net";
import { dirname, join } from "node:path";

/**
 * Bridgens identitet pr. bruger (#11): adgangstoken og port.
 *
 * På delte maskiner (Remote Desktop, Citrix, terminalservere) kan alle brugere nå 127.0.0.1.
 * Host-, Origin- og X-MIAI-tjekket holder fremmede websider ude, men ikke andre lokale processer.
 * Derfor kræver API'et et token, som kun brugeren selv kender:
 *
 *  - tokenet gemmes i brugerens konfigurationsmappe (bridge.json, kun læsbar af brugeren)
 *  - bridgen skriver et brugerspecifikt manifest, hvor task pane-adressen indeholder tokenet
 *    (…/taskpane.html?token=…); manifestet ligger også i brugerens egen mappe
 *  - add-in'et læser tokenet fra sin egen adresse og sender det med hvert API-kald
 *
 * Den serverede taskpane.html indeholder ikke tokenet – så kan en anden proces ikke bare hente det.
 */

export const DEFAULT_PORT = 44321;
/** Porte, der prøves, hvis standardporten er optaget (fx af en anden brugers bridge). */
const PORT_RANGE = 80;

export interface BridgeIdentity {
  token: string;
  port: number;
}

interface StoredIdentity {
  token?: string;
  port?: number;
}

export function identityFile(configDir: string): string {
  return join(configDir, "bridge.json");
}

function readStored(file: string): StoredIdentity {
  try {
    return existsSync(file) ? (JSON.parse(readFileSync(file, "utf8")) as StoredIdentity) : {};
  } catch {
    return {};
  }
}

function writeStored(file: string, data: StoredIdentity): void {
  mkdirSync(dirname(file), { recursive: true });
  const tmp = `${file}.tmp`;
  writeFileSync(tmp, JSON.stringify(data, null, 2), { mode: 0o600 });
  renameSync(tmp, file);
  try {
    chmodSync(file, 0o600);
  } catch {
    // Windows: filen arver brugerprofilens adgangsrettigheder.
  }
}

export function portIsFree(port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const s = createServer();
    s.once("error", () => resolve(false));
    s.listen(port, "127.0.0.1", () => s.close(() => resolve(true)));
  });
}

/**
 * Henter (eller opretter) brugerens token og port. Porten huskes, så manifestet forbliver
 * gyldigt; er den gemte port optaget, vælges den næste ledige fra standardporten og frem.
 * `fixedPort` (MIAI_PORT) bruges uændret.
 */
export async function loadIdentity(configDir: string, fixedPort?: number): Promise<BridgeIdentity> {
  const file = identityFile(configDir);
  const stored = readStored(file);
  const token = stored.token && stored.token.length >= 32 ? stored.token : randomBytes(32).toString("base64url");

  let port = fixedPort;
  if (!port) {
    const candidates = [stored.port, ...Array.from({ length: PORT_RANGE }, (_, i) => DEFAULT_PORT + i)].filter((p): p is number => !!p);
    for (const p of candidates) {
      if (await portIsFree(p)) {
        port = p;
        break;
      }
    }
    if (!port) throw new Error(`Ingen ledig port mellem ${DEFAULT_PORT} og ${DEFAULT_PORT + PORT_RANGE - 1}.`);
  }

  if (stored.token !== token || stored.port !== port) writeStored(file, { token, port });
  return { token, port };
}

/** Sammenligner tokens i konstant tid. */
export function tokenMatches(expected: string, given: unknown): boolean {
  if (typeof given !== "string") return false;
  const a = Buffer.from(expected);
  const b = Buffer.from(given);
  return a.length === b.length && timingSafeEqual(a, b);
}

/**
 * Laver brugerens manifest ud fra bridge-manifestet (https://localhost:44321): den faktiske port
 * indsættes, og task pane-adressen får tokenet som parameter.
 */
export function personalizeManifest(baseXml: string, id: BridgeIdentity): string {
  const origin = `https://localhost:${id.port}`;
  return baseXml.replaceAll("https://localhost:44321", origin).replaceAll(`${origin}/taskpane.html"`, `${origin}/taskpane.html?token=${id.token}"`);
}

/** Skriver brugerens manifest i konfigurationsmappen (kun læsbart af brugeren) og returnerer stien. */
export function writePersonalManifest(configDir: string, baseManifestFile: string, id: BridgeIdentity): string | undefined {
  if (!existsSync(baseManifestFile)) return undefined;
  const out = join(configDir, "manifest.xml");
  const xml = personalizeManifest(readFileSync(baseManifestFile, "utf8"), id);
  const current = existsSync(out) ? readFileSync(out, "utf8") : "";
  if (current !== xml) {
    mkdirSync(configDir, { recursive: true });
    writeFileSync(out, xml, { mode: 0o600 });
    try {
      chmodSync(out, 0o600);
    } catch {
      // Windows
    }
  }
  return out;
}
