// Automatiske opdateringer (#5, ADR-0016) via electron-updater: henter nye udgaver fra
// opdateringsfeedet (build.publish i package.json – en "generic" URL på GitHub Pages) i
// baggrunden og installerer dem, når appen genstartes. macOS kræver en signeret app;
// uden signering logges fejlen, og appen virker som før.
import { app } from "electron";

/** Tjek ved start og derefter hver 6. time. */
const INTERVAL_MS = 6 * 60 * 60 * 1000;

let ready = null; // version, der er hentet og klar til installation
let updater = null;

export async function setupUpdates({ log, onChange }) {
  if (!app.isPackaged || process.env.MIAI_NO_UPDATES) return;
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
