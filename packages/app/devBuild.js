// Udviklerudgaven (ADR-0016, tillæg): hvert push til main bygger en .dmg til Apple Silicon, som
// ligger på GitHub Pages under /dev/ sammen med dev.json. Her er den del, der ikke kræver Electron:
// at læse dev.json sikkert og afgøre, hvad menupunktet skal vise.

/** Hvor udviklerudgaven ligger. MIAI_DEV_FEED kan pege et andet sted hen (test, egen server). */
export const DEV_FEED = "https://mor10w.github.io/MIAIPlugin/dev/";

const VERSION = /^\d+\.\d+\.\d+-dev\.(\d+)$/;
const FILE = /^[A-Za-z0-9][A-Za-z0-9._-]*\.dmg$/;

/** Udviklerudgaven bygges indtil videre kun til Apple Silicon (Windows: #40). */
export function devBuildSupported(platform, arch) {
  return platform === "darwin" && arch === "arm64";
}

/** Byggenummeret i en udviklerversion ("0.1.0-dev.42" → 42), eller null for en almindelig version. */
export function devBuildNumber(version) {
  const m = VERSION.exec(String(version ?? ""));
  return m ? Number(m[1]) : null;
}

/**
 * Fortolker dev.json. Kun en udviklerversion og et rent filnavn accepteres, så feedet aldrig kan
 * sende brugeren til en anden adresse end mappen selv. Returnerer null, hvis indholdet ikke passer.
 */
export function parseDevInfo(json, feed = DEV_FEED) {
  if (!json || typeof json !== "object") return null;
  const { version, file, commit, date } = json;
  if (devBuildNumber(version) === null || typeof file !== "string" || !FILE.test(file)) return null;
  const base = new URL(feed);
  if (base.protocol !== "https:") return null;
  return {
    version,
    url: new URL(file, base).href,
    commit: typeof commit === "string" ? commit.slice(0, 12) : undefined,
    date: typeof date === "string" ? date.slice(0, 10) : undefined,
  };
}

/** Henter dev.json fra feedet. Returnerer null ved enhver fejl (ingen udviklerudgave, intet net). */
export async function fetchDevInfo(feed = DEV_FEED, fetchImpl = fetch) {
  try {
    const res = await fetchImpl(new URL("dev.json", feed), { cache: "no-store", signal: AbortSignal.timeout(10_000) });
    if (!res.ok) return null;
    return parseDevInfo(await res.json(), feed);
  } catch {
    return null;
  }
}

/**
 * Hvad menupunktet viser: den nyeste udviklerudgave kan hentes, den er allerede installeret, eller
 * (når feedet ikke kan nås) intet punkt. `current` er appens egen version.
 */
export function devMenuState(current, info) {
  if (!info) return { kind: "none" };
  if (info.version === current) return { kind: "current", label: `Udviklerudgave ${info.version} er installeret` };
  const mine = devBuildNumber(current);
  const theirs = devBuildNumber(info.version);
  // En ældre udviklerudgave på feedet end den installerede tilbydes ikke.
  if (mine !== null && theirs !== null && theirs < mine) return { kind: "current", label: `Udviklerudgave ${current} er installeret` };
  return { kind: "available", label: `Hent udviklerudgave ${info.version}${info.date ? ` (${info.date})` : ""}…`, url: info.url };
}
