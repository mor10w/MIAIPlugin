// Automatiske opdateringer (#5, ADR-0016) via electron-updater: henter nye udgaver fra
// opdateringsfeedet (build.publish i package.json – en "generic" URL på GitHub Pages) i
// baggrunden og installerer dem, når appen genstartes. macOS kræver en signeret app;
// uden signering logges fejlen, og appen virker som før.
import { app, net, shell } from "electron";
import { DEV_FEED, devBuildNumber, devBuildSupported, devMenuState, fetchDevInfo, platformKey } from "./devBuild.js";

/** Tjek ved start og derefter hver 6. time. */
const INTERVAL_MS = 6 * 60 * 60 * 1000;

let ready = null; // version, der er hentet og klar til installation
let updater = null;

export async function setupUpdates({ log, onChange }) {
  if (!app.isPackaged || process.env.MIAI_NO_UPDATES) return;
  // En udviklerudgave følger udviklerkanalen (menupunktet nedenfor), ikke de udgivne versioner.
  if (devBuildNumber(app.getVersion()) !== null) return;
  try {
    // CommonJS-pakke: default-eksporten indeholder autoUpdater.
    const mod = await import("electron-updater");
    updater = (mod.default ?? mod).autoUpdater;
  } catch (err) {
    log(`electron-updater kunne ikke indlæses: ${err.message}`);
    return;
  }
  updater.autoDownload = true;
  updater.autoInstallOnAppQuit = true;
  updater.logger = { info: (m) => log(`opdatering: ${m}`), warn: (m) => log(`opdatering: ${m}`), error: (m) => log(`opdatering FEJL: ${m}`), debug: () => {} };
  updater.on("update-downloaded", (info) => {
    ready = info.version;
    log(`opdatering ${info.version} hentet – installeres ved genstart`);
    onChange();
  });
  updater.on("error", (err) => log(`opdatering FEJL: ${err?.message ?? err}`));
  const check = () => updater.checkForUpdates().catch((err) => log(`opdateringstjek fejlede: ${err.message}`));
  void check();
  setInterval(check, INTERVAL_MS).unref?.();
}

/** Menupunkt, når en opdatering er klar ("Genstart og opdater til 0.2.0"). */
export function updateMenuItem() {
  if (!ready || !updater) return [];
  return [{ label: `Genstart og opdater til ${ready}`, click: () => updater.quitAndInstall() }];
}

// --- Udviklerudgaven (ADR-0016, tillæg) --------------------------------------------------------
// Hvert push til main bygger en .dmg til Apple Silicon og en installer til Windows (x64, #40).
// Appen installerer den ikke selv (på Mac kræver det en signeret app); menupunktet åbner hentningen
// i browseren. På Mac trækkes appen over som ved første installation; på Windows køres installeren,
// som installerer oven i den nuværende udgave.

let devInfo = null; // seneste dev.json fra feedet

export function setupDevBuilds({ log, onChange }) {
  if (!app.isPackaged || process.env.MIAI_NO_UPDATES) return;
  if (!devBuildSupported(process.platform, process.arch)) return;
  const feed = process.env.MIAI_DEV_FEED || DEV_FEED;
  // net.fetch følger systemets proxyindstillinger, det gør Nodes fetch ikke.
  const check = async () => {
    const info = await fetchDevInfo(feed, (url, init) => net.fetch(String(url), init), platformKey(process.platform, process.arch));
    if (info?.version === devInfo?.version) return;
    devInfo = info;
    if (info) log(`udviklerudgave ${info.version} findes på ${feed}`);
    onChange();
  };
  void check();
  setInterval(() => void check(), INTERVAL_MS).unref?.();
}

/** Menupunkt for udviklerudgaven: "Hent udviklerudgave …", eller at den nyeste er installeret. */
export function devMenuItem() {
  const state = devMenuState(app.getVersion(), devInfo);
  if (state.kind === "none") return [];
  if (state.kind === "current") return [{ label: state.label, enabled: false }];
  return [{ label: state.label, click: () => void shell.openExternal(state.url) }];
}
