// Opdatering fra menuen (ADR-0016): den del, der ikke kræver Electron. Hvordan appen opdateres på
// hver platform, hvad menuen viser i hver tilstand, og hvornår brugeren får en besked.

import { devBuildNumber } from "./devBuild.js";

/** Opdateringsfeedet for de almindelige udgivelser (build.publish i package.json). */
export const UPDATE_FEED = "https://mor10w.github.io/MIAIPlugin/updates/";

/**
 * Hvordan appen opdateres:
 * - "auto": electron-updater henter i baggrunden og installerer ved genstart (Windows).
 * - "auto-dev": som "auto", men fra udviklerkanalen (dev.yml under /dev/); kun Windows-udviklerudgaven.
 * - "notify": appen siger til, når der er en ny udgave, og hentningen åbnes i browseren. Mac, indtil
 *   appen er signeret med Developer ID (#67): Squirrel.Mac afviser at installere en ad hoc-signeret app.
 * - "browser-dev": Mac-udviklerudgaven; dev.json læses, og hentningen åbnes i browseren.
 * - "none": ingen opdatering (fx Linux).
 */
export function updateMode(platform, version) {
  const dev = devBuildNumber(version) !== null;
  if (platform === "win32") return dev ? "auto-dev" : "auto";
  if (platform === "darwin") return dev ? "browser-dev" : "notify";
  return "none";
}

const RELEASE = /^\d+\.\d+\.\d+$/;

/** Adressen på Mac-dmg'en for en udgivelse på feedet (artifactName: MIAI-<version>-<arch>.dmg), eller null. */
export function macDownloadUrl(version, arch, feed = UPDATE_FEED) {
  if (!RELEASE.test(String(version ?? "")) || (arch !== "arm64" && arch !== "x64")) return null;
  const base = new URL(feed);
  if (base.protocol !== "https:") return null;
  return new URL(`MIAI-${version}-${arch}.dmg`, base).href;
}

/**
 * Menupunkterne for opdatering i en given tilstand. `state.status` er en af:
 * idle, checking, latest, error, available (kan hentes i browseren), downloading, ready (klar til genstart).
 * Returnerer beskrivelser: { label, action?: "check" | "install" | "open", url?, enabled? }.
 */
export function updateMenuItems(state) {
  const v = state?.version;
  switch (state?.status) {
    case "checking":
      return [{ label: "Søger efter opdateringer…", enabled: false }];
    case "downloading":
      return [{ label: `Henter MIAI ${v}…`, enabled: false }];
    case "ready":
      return [{ label: `Genstart og opdater til ${v}`, action: "install" }];
    case "available":
      return [{ label: `Hent MIAI ${v}…`, action: "open", url: state.url }];
    default:
      return [{ label: "Søg efter opdateringer…", action: "check" }];
  }
}

/**
 * Beskeden (systemnotifikation) for en ny tilstand, eller null. Om en ny udgave siges der til én gang
 * pr. version, også når appen selv søgte i baggrunden. "Nyeste version" og fejl vises kun, når
 * brugeren selv bad om at søge (`manual`), så baggrundstjekket aldrig forstyrrer.
 */
export function updateNotification(state, { manual = false, current, notified = new Set() } = {}) {
  const v = state?.version;
  switch (state?.status) {
    case "ready":
      if (notified.has(`ready:${v}`)) return null;
      return {
        key: `ready:${v}`,
        title: "MIAI er klar til at opdatere",
        body: `Version ${v} er hentet. Klik for at genstarte og opdatere.`,
        action: "install",
      };
    case "available":
      if (notified.has(`available:${v}`)) return null;
      return { key: `available:${v}`, title: "Ny version af MIAI", body: `Version ${v} er udgivet. Klik for at hente den.`, action: "open", url: state.url };
    case "downloading":
      return manual ? { title: "MIAI opdateres", body: `Version ${v} hentes i baggrunden. Du får besked, når den er klar.` } : null;
    case "latest":
      return manual ? { title: "MIAI er opdateret", body: `Du har den nyeste version${current ? ` (${current})` : ""}.` } : null;
    case "error":
      return manual ? { title: "Kunne ikke søge efter opdateringer", body: state.error || "Prøv igen senere." } : null;
    default:
      return null;
  }
}
