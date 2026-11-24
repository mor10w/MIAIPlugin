// Opdateringer (#5, ADR-0016). Menuen har altid "Søg efter opdateringer…", og appen giver besked
// (systemnotifikation), når en ny udgave er klar. Hvordan der opdateres, afhænger af platformen
// (updateMode i updateState.js):
// - Windows: electron-updater henter i baggrunden fra feedet på GitHub Pages og installerer ved
//   genstart eller fra menuen. Udviklerudgaven følger sin egen kanal (dev.yml under /dev/).
// - Mac: Squirrel.Mac kræver en Developer ID-signeret app (#67). Indtil da siger appen til, når der
//   er en ny udgave, og åbner hentningen i browseren.
import { app, net, Notification, shell } from "electron";
import { DEV_FEED, devBuildNumber, devBuildSupported, devMenuState, fetchDevInfo, platformKey } from "./devBuild.js";
import { macDownloadUrl, updateMenuItems, updateMode, updateNotification } from "./updateState.js";

/** Tjek ved start og derefter hver 6. time. */
const INTERVAL_MS = 6 * 60 * 60 * 1000;

let mode = "none";
let updater = null;
let state = { status: "idle" };
/** Brugeren bad selv om at søge: så siges der også til, når der ingen opdatering er, eller det fejlede. */
let manual = false;
/** Versioner, der allerede er givet besked om (én besked pr. version). */
const notified = new Set();
/** Notifikationer holdes i live, til de lukkes; ellers kan klikket gå tabt. */
const shown = new Set();
let log = () => {};
let onChange = () => {};

function notify(n) {
  if (n.key) notified.add(n.key);
  if (!Notification.isSupported()) return;
  const note = new Notification({ title: n.title, body: n.body });
  shown.add(note);
  note.on("close", () => shown.delete(note));
  if (n.action) note.on("click", () => act(n));
  note.show();
}

function setState(next) {
  state = next;
  const n = updateNotification(state, { manual, current: app.getVersion(), notified });
  if (n) notify(n);
  // Søgningen, brugeren startede, er afsluttet (en hentning fortsætter i baggrunden og giver selv besked).
  if (state.status !== "checking" && state.status !== "downloading") manual = false;
  onChange();
}

function act(item) {
  if (item.action === "install") updater?.quitAndInstall();
  else if (item.action === "open" && item.url) void shell.openExternal(item.url);
  else if (item.action === "check") checkNow();
}

function check() {
  updater?.checkForUpdates().catch((err) => {
    log(`opdateringstjek fejlede: ${err.message}`);
    if (state.status !== "ready") setState({ status: "error", error: err.message });
  });
}

/** "Søg efter opdateringer…" i menuen. */
export function checkNow() {
  manual = true;
  if (mode === "browser-dev") void checkDev();
  else if (updater) check();
  else manual = false;
}

export async function setupUpdates(opts) {
  ({ log, onChange } = opts);
  if (!app.isPackaged || process.env.MIAI_NO_UPDATES) return;
  mode = updateMode(process.platform, app.getVersion());
  // Mac-udviklerudgaven følger dev.json (setupDevBuilds nedenfor).
  if (mode === "none" || mode === "browser-dev") return;
  try {
    // CommonJS-pakke: default-eksporten indeholder autoUpdater.
    const mod = await import("electron-updater");
    updater = (mod.default ?? mod).autoUpdater;
  } catch (err) {
    log(`electron-updater kunne ikke indlæses: ${err.message}`);
    return;
  }
  // Mac: kun besked; Squirrel.Mac kan ikke installere en ad hoc-signeret app (#67).
  updater.autoDownload = mode !== "notify";
  updater.autoInstallOnAppQuit = mode !== "notify";
  if (mode === "auto-dev") {
    // Udviklerkanalen: dev.yml i sin egen mappe, så almindelige brugere aldrig får et udviklerbyg.
    // Kanalen sættes i feedet (ikke updater.channel, som også ville tillade nedgradering).
    updater.setFeedURL({ provider: "generic", url: process.env.MIAI_DEV_FEED || DEV_FEED, channel: "dev" });
  }
  updater.logger = { info: (m) => log(`opdatering: ${m}`), warn: (m) => log(`opdatering: ${m}`), error: (m) => log(`opdatering FEJL: ${m}`), debug: () => {} };
  updater.on("checking-for-update", () => {
    if (state.status !== "ready" && state.status !== "downloading") setState({ status: "checking" });
  });
  updater.on("update-not-available", () => {
    if (state.status !== "ready") setState({ status: "latest" });
  });
  updater.on("update-available", (info) => {
    // Allerede hentet og klar (et nyt tjek melder den samme igen): bliv ved "Genstart og opdater".
    if (state.status === "ready" && state.version === info.version) {
      if (manual) notify(updateNotification(state, { notified: new Set() }));
      manual = false;
      return onChange();
    }
    if (mode !== "notify") return setState({ status: "downloading", version: info.version });
    const url = macDownloadUrl(info.version, process.arch);
    if (url) setState({ status: "available", version: info.version, url });
    else setState({ status: "error", error: `Ukendt udgave ${info.version}` });
  });
  updater.on("update-downloaded", (info) => {
    log(`opdatering ${info.version} hentet – installeres ved genstart`);
    setState({ status: "ready", version: info.version });
  });
  updater.on("error", (err) => {
    log(`opdatering FEJL: ${err?.message ?? err}`);
    if (state.status !== "ready") setState({ status: "error", error: err?.message ?? String(err) });
  });
  check();
  setInterval(check, INTERVAL_MS).unref?.();
}

/** Opdateringens menupunkt: "Søg efter opdateringer…", "Henter …", "Genstart og opdater til …" eller "Hent MIAI …". */
export function updateMenuItem() {
  if (mode === "none") return [];
  return updateMenuItems(state).map((item) => (item.action ? { label: item.label, click: () => act(item) } : { label: item.label, enabled: false }));
}

// --- Udviklerudgaven (ADR-0016, tillæg) --------------------------------------------------------
// Hvert push til main bygger en .dmg til Apple Silicon og en installer til Windows (x64, #40).
// En almindelig udgave viser "Hent udviklerudgave …", når der findes en; klikket åbner hentningen i
// browseren. En installeret Windows-udviklerudgave opdaterer sig selv (kanalen "dev" ovenfor). Mac-
// udviklerudgaven får besked og åbner hentningen i browseren (Squirrel.Mac kræver signering).

let devInfo = null; // seneste dev.json fra feedet
let devFeed = DEV_FEED;

async function checkDev() {
  // net.fetch følger systemets proxyindstillinger, det gør Nodes fetch ikke.
  const info = await fetchDevInfo(devFeed, (url, init) => net.fetch(String(url), init), platformKey(process.platform, process.arch));
  const changed = info?.version !== devInfo?.version;
  devInfo = info;
  if (info && changed) log(`udviklerudgave ${info.version} findes på ${devFeed}`);
  // Kun en installeret udviklerudgave får besked; en almindelig udgave ser blot menupunktet.
  if (mode === "browser-dev") {
    const s = devMenuState(app.getVersion(), devInfo);
    if (s.kind === "available") setState({ status: "available", version: info.version, url: s.url });
    else if (!info) setState({ status: "error", error: "Udviklerudgaven kunne ikke hentes." });
    else setState({ status: "latest" });
  } else if (changed) onChange();
}

export function setupDevBuilds(opts) {
  ({ log, onChange } = opts);
  if (!app.isPackaged || process.env.MIAI_NO_UPDATES) return;
  if (!devBuildSupported(process.platform, process.arch)) return;
  // Windows-udviklerudgaven opdaterer sig selv via electron-updater.
  if (devBuildNumber(app.getVersion()) !== null && process.platform === "win32") return;
  devFeed = process.env.MIAI_DEV_FEED || DEV_FEED;
  void checkDev();
  setInterval(() => void checkDev(), INTERVAL_MS).unref?.();
}

/** Menupunkt for udviklerudgaven: "Hent udviklerudgave …", eller at den nyeste er installeret. */
export function devMenuItem() {
  // Mac-udviklerudgaven: en nyere vises allerede som "Hent MIAI …" i opdateringspunktet ovenfor.
  if (mode === "browser-dev" && state.status === "available") return [];
  const s = devMenuState(app.getVersion(), devInfo);
  if (s.kind === "none") return [];
  if (s.kind === "current") return [{ label: s.label, enabled: false }];
  return [{ label: s.label, click: () => void shell.openExternal(s.url) }];
}
