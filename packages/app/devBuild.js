// Udviklerudgaven (ADR-0016, tillæg): hvert push til main bygger en .dmg til Apple Silicon og en
// installer til Windows (x64, #40), som ligger på GitHub Pages under /dev/ sammen med dev.json.
// Her er den del, der ikke kræver Electron: at læse dev.json sikkert og afgøre, hvad menupunktet
// skal vise.

/** Hvor udviklerudgaven ligger. MIAI_DEV_FEED kan pege et andet sted hen (test, egen server). */
export const DEV_FEED = "https://mor10w.github.io/MIAIPlugin/dev/";

const VERSION = /^\d+\.\d+\.\d+-dev\.(\d+)$/;

/** Platformene, udviklerudgaven bygges til, og den filtype, hver af dem må hente. */
const PLATFORMS = {
  "darwin-arm64": /^[A-Za-z0-9][A-Za-z0-9._-]*\.dmg$/,
  "win32-x64": /^[A-Za-z0-9][A-Za-z0-9._-]*\.exe$/,
};

/** Nøglen i dev.json's `files` for en platform og arkitektur, fx "win32-x64". */
export function platformKey(platform, arch) {
  return `${platform}-${arch}`;
}

/** Udviklerudgaven bygges til Mac på Apple Silicon og Windows på x64. */
export function devBuildSupported(platform, arch) {
  return Object.hasOwn(PLATFORMS, platformKey(platform, arch));
}

/** Byggenummeret i en udviklerversion ("0.1.0-dev.42" → 42), eller null for en almindelig version. */
export function devBuildNumber(version) {
  const m = VERSION.exec(String(version ?? ""));
  return m ? Number(m[1]) : null;
}

/**
 * Fortolker dev.json for én platform. Kun en udviklerversion og et rent filnavn med platformens
 * filtype accepteres, så feedet aldrig kan sende brugeren til en anden adresse end mappen selv.
 * Returnerer null, hvis indholdet ikke passer, eller der ikke er et byg til platformen.
 *
 * `files` har ét filnavn pr. platform (#40). Ældre feeds har kun `file` (Mac-dmg'en); det læses
 * stadig for Mac, og nye feeds skriver det også, så ældre udviklerudgaver kan finde opdateringen.
 */
export function parseDevInfo(json, feed = DEV_FEED, key = "darwin-arm64") {
  if (!json || typeof json !== "object") return null;
  const { version, files, commit, date } = json;
  const pattern = Object.hasOwn(PLATFORMS, key) ? PLATFORMS[key] : undefined;
  if (!pattern || devBuildNumber(version) === null) return null;
  const file = files && typeof files === "object" ? files[key] : key === "darwin-arm64" ? json.file : undefined;
  if (typeof file !== "string" || !pattern.test(file)) return null;
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
export async function fetchDevInfo(feed = DEV_FEED, fetchImpl = fetch, key = "darwin-arm64") {
  try {
    const res = await fetchImpl(new URL("dev.json", feed), { cache: "no-store", signal: AbortSignal.timeout(10_000) });
    if (!res.ok) return null;
    return parseDevInfo(await res.json(), feed, key);
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
