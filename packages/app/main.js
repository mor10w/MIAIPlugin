// MIAI App – menu-bar-app, der kører MIAI Bridge på https://localhost:44321.
// Salgsudgave B (ADR-0008): alt kører lokalt; ingen vendor-cloud, ingen CORS.
// Bridgen (Fastify) kører i en Electron utility-process – brugeren behøver ikke Node.js.
// VIGTIGT: bridgen kan IKKE køre i main-processen – Chromiums netværksspecifik forstyrrer
// Node's TLS-server, så håndskaber hænger. I en utility-process (uden Chromium) fungerer
// TLS som i almindelig Node – verificeret med ELECTRON_RUN_AS_NODE + curl.
import { app, BrowserWindow, dialog, Menu, Tray, nativeImage, shell, utilityProcess } from "electron";
import { appendFileSync, existsSync, unlinkSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { VERSION as BRIDGE_VERSION } from "@miai/bridge/dist/server.js";
import { configDir } from "@miai/bridge/dist/config.js";
import { currentPlatform } from "./platform.js";
import { devMenuItem, setupDevBuilds, setupUpdates, updateMenuItem } from "./updates.js";
import { launchInfo, secondInstanceDialog } from "./instance.js";

/** Den faktiske port meldes af bridge-processen (pr. bruger, #11). */
let PORT = Number(process.env.MIAI_PORT ?? 44321);
/** Brugerens personlige manifest (med port og adgangstoken), skrevet af bridge-processen. */
let PERSONAL_MANIFEST = null;
/** Samme mappe som bridgen bruger (macOS: ~/Library/Application Support/MIAI, Windows: %APPDATA%\\MIAI). */
const MIAI_DIR = configDir();
const CERT_FILE = join(MIAI_DIR, "tls", "localhost.crt");

/** Fillog: Electron svulger console-output i pakker, så fejl kan ellers ikke ses. */
function log(msg) {
  try {
    appendFileSync(join(MIAI_DIR, "MIAI.log"), `[${new Date().toISOString()}] ${msg}\n`);
  } catch {
    /* log skal aldrig fælde appen */
  }
}

// Packaget: filer ligger i Resources/ (electron-builder extraResources). I dev: repoet.
// main.js ligger i packages/app, så søsterpakken er ../addin.
const ADDIN_DIR = app.isPackaged ? join(process.resourcesPath, "addin") : fileURLToPath(new URL("../addin/dist", import.meta.url));
/** Bridge-manifestet, som det personlige manifest laves ud fra. */
const MANIFEST = app.isPackaged ? join(process.resourcesPath, "manifest.xml") : fileURLToPath(new URL("../addin/manifest.xml", import.meta.url));
/** Outlooks eget manifest (MailApp, #7) – installeres manuelt i Outlook (se BRUG.md). */
const OUTLOOK_MANIFEST = app.isPackaged
  ? join(process.resourcesPath, "manifest.outlook.xml")
  : fileURLToPath(new URL("../addin/manifest.outlook.xml", import.meta.url));

let bridgeProc = null;
let tray = null;
let quitting = false;
const bridgeState = { listening: false };

// --- Bridgen (utility-process) -------------------------------------------------
/**
 * Fork'er bridgen som utility-process og afventer, at den lytter (eller fejler).
 * Returnerer en Promise, der løses senest efter 10 s.
 */
function startBridge() {
  log("startBridge: fork'er utility-process");
  bridgeProc = utilityProcess.fork(
    join(app.getAppPath(), "bridge-process.mjs"),
    [],
    // MIAI_PORT sendes kun videre, hvis brugeren selv har sat den; ellers vælger bridgen en ledig port pr. bruger.
    { env: { ...process.env, MIAI_ADDIN_DIR: ADDIN_DIR, MIAI_MANIFEST: MANIFEST, MIAI_OUTLOOK_MANIFEST: OUTLOOK_MANIFEST } },
  );

  bridgeProc.on("message", (msg) => {
    if (msg?.type === "listening") {
      bridgeState.listening = true;
      if (msg.port) PORT = msg.port;
      if (msg.manifest) PERSONAL_MANIFEST = msg.manifest;
      log(`bridge-process: lytter på ${msg.addr} (v${msg.version})`);
      setStatus("Bridgen kører");
    } else if (msg?.type === "status") {
      setStatus(statusFromEndpoint(msg.endpoint));
    } else if (msg?.type === "error") {
      log(`bridge-process FEJL: ${msg.message}`);
      quitting = true; // undgå en ekstra "stoppede uventet"-dialog, når processen afslutter
      dialog.showErrorBox("MIAI", `Bridgen kunne ikke startes: ${msg.message}`);
      app.quit();
    }
  });

  bridgeProc.on("exit", (code) => {
    log(`bridge-process afsluttet med kode ${code}`);
    if (!quitting && code !== 0) {
      dialog.showErrorBox("MIAI", `Bridgen stoppede uventet (kode ${code}). Se MIAI.log for detaljer.`);
      app.quit();
    }
  });

  return new Promise((resolve) => {
    const t = setTimeout(resolve, 10_000);
    bridgeProc.on("message", (m) => {
      if (m?.type === "listening") {
        clearTimeout(t);
        resolve();
      }
    });
    bridgeProc.on("exit", () => {
      clearTimeout(t);
      resolve();
    });
  });
}

// --- Office og certifikat (pr. platform – platform.js) --------------------------
const platform = currentPlatform();

/** Installerer brugerens personlige manifest i Word, Excel og PowerPoint. */
function installInOffice() {
  if (!PERSONAL_MANIFEST) throw new Error("Bridgen har ikke skrevet dit manifest endnu.");
  platform.installInOffice(PERSONAL_MANIFEST);
}

/** ok · stale (anden port/token) · missing */
function officeInstallState() {
  return PERSONAL_MANIFEST ? platform.officeInstallState(PERSONAL_MANIFEST) : "missing";
}

function trustCert() {
  return platform.trustCert(CERT_FILE);
}

function isCertTrusted() {
  return platform.isCertTrusted(CERT_FILE);
}

// --- Dialoger -----------------------------------------------------------------------
// En menulinje-app er ikke i forgrunden, så en dialog uden vindue kan gemme sig bag andre
// vinduer. Og mens en sådan (app-modal) dialog er åben, er hovedtråden optaget: startes MIAI
// igen, venter den nye i 20 s på svar (Dock-ikonet hopper) og dræber så den kørende
// (Chromiums enkelt-instans-lås). Set på macOS i praksis. Derfor vises dialogerne som et ark på
// et lille, synligt vindue: dialogen kommer frem, og appen svarer imens.
const DIALOG_PAGE = `data:text/html;charset=utf-8,${encodeURIComponent(
  '<!doctype html><meta charset="utf-8"><title>MIAI</title><body style="margin:0;background:#f6f7f9"></body>',
)}`;

/** Som dialog.showMessageBox, men synlig og uden at låse appen. */
async function ask(options) {
  const win = new BrowserWindow({
    width: 520,
    height: 300,
    show: false,
    title: options.title ?? "MIAI",
    resizable: false,
    minimizable: false,
    maximizable: false,
    fullscreenable: false,
    backgroundColor: "#f6f7f9",
    webPreferences: { sandbox: true, contextIsolation: true, javascript: false },
  });
  try {
    await win.loadURL(DIALOG_PAGE);
    win.center();
    win.show();
    app.focus({ steal: true });
    win.focus();
    return await dialog.showMessageBox(win, options);
  } finally {
    if (!win.isDestroyed()) win.destroy();
  }
}

/** Viser install-fejl. EPERM = macOS blokerede adgangen (TCC) – der gives en vej til Full Disk Access. */
function showInstallError(err, lead) {
  const prefix = lead ? `${lead}\n\n` : "";
  if (err?.code === "EPERM" && process.platform === "darwin") {
    // I dev er det Electron-programmet selv, der skal have adgangen.
    const target = app.isPackaged ? "MIAI" : `Electron (${process.execPath})`;
    void ask({
      type: "warning",
      title: "Adgang nægtet af macOS",
      message: "macOS blokerede MIAI fra at skrive i Office-appernes mapper.",
      detail: `${prefix}Giv ${target} fuld diskadgang under Systemindstillinger → Anonymitet og sikkerhed → Fuld diskadgang (fjern en ældre MIAI fra listen), og vælg så Installer i Office… i MIAI-menuen.`,
      buttons: ["Åbn Indstillinger", "Luk"],
      defaultId: 0,
    }).then((r) => {
      if (r.response === 0) void shell.openExternal("x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles");
    });
    return;
  }
  void ask({ type: "error", title: "MIAI", message: "Kunne ikke installere i Office", detail: `${prefix}${err?.message ?? String(err)}`, buttons: ["OK"] });
}

function showCertResult(r) {
  return ask({
    type: r.ok ? "info" : "warning",
    title: "Certifikat",
    message: r.ok ? "MIAI-certifikatet er nu betroet." : `Certifikatet blev ikke betroet: ${r.message || "ukendt fejl"}`,
    detail: r.ok ? "Office kan nu forbinde til MIAI Bridge." : "Prøv igen, eller genstart appen.",
  });
}

// --- Menu --------------------------------------------------------------------
let statusText = "Starter…";

function buildMenu() {
  return Menu.buildFromTemplate([
    { label: `MIAI Bridge v${BRIDGE_VERSION}`, enabled: false },
    { label: statusText, enabled: false },
    { type: "separator" },
    {
      label: "Installer i Office…",
      click: () => {
        try {
          installInOffice();
          ask({
            type: "info",
            title: "MIAI er installeret i Office",
            message: "Manifestet er kopieret til Word, Excel og PowerPoint.",
            detail: "Genstart de Office-apps, du vil bruge MIAI i. Siden da finder du MIAI under Indsæt → Tilføjelser (og som 'AI-assistent' på fanen Hjem).",
          });
        } catch (err) {
          showInstallError(err);
        }
      },
    },
    {
      label: "Betro certifikat…",
      click: () => showCertResult(trustCert()),
    },
    { label: `Åbn https://localhost:${PORT}`, click: () => void shell.openExternal(`https://localhost:${PORT}`) },
    {
      // Outlook kan ikke installeres automatisk som Word/Excel/PowerPoint: brugeren tilføjer filen i Outlook.
      label: `Vis Outlook-manifest i ${platform.fileManager}…`,
      click: () => shell.showItemInFolder(join(MIAI_DIR, "manifest-outlook.xml")),
    },
    {
      type: "checkbox",
      label: app.isPackaged ? "Start automatisk ved login" : "Start automatisk ved login (kun i den installerede app)",
      enabled: app.isPackaged,
      checked: startAtLoginEnabled(),
      click: (item) => void setStartAtLogin(item.checked),
    },
    ...updateMenuItem(),
    ...devMenuItem(),
    { type: "separator" },
    { label: "Afslut MIAI", click: () => app.quit() },
  ]);
}

/** Bygger menuen igen (fx når login-elementets tilstand er ændret). */
function refreshMenu() {
  if (tray) tray.setContextMenu(buildMenu());
}

let lastStatus = "";
/** Opdaterer menuens statuslinje (kaldes fra bridge-processens messages). */
function setStatus(s) {
  if (s === lastStatus) return;
  log(`status: ${s}`);
  lastStatus = s;
  statusText = s;
  if (tray) {
    tray.setToolTip(statusText);
    tray.setContextMenu(buildMenu());
  }
}

function statusFromEndpoint(ep) {
  if (!ep) return bridgeState.listening ? "Bridgen kører" : "Bridgen svarer ikke";
  if (ep.reachable) return `Forbundet · ${ep.baseUrl}`;
  return `Modellen svarer ikke (${ep.error ?? "ukendt fejl"})`;
}

// --- Start ved login ------------------------------------------------------------
// macOS' egen login-element-API via Electron (SMAppService på macOS 13+): elementet vises under
// Systemindstillinger → Generelt → Login-emner, følger app-bundlen, hvis den flyttes, og kan
// kræve brugerens godkendelse. Kun i den installerede app – i dev ville det starte en tom Electron.

/** LaunchAgent fra en tidligere udgave af denne funktion – fjernes, så appen ikke startes to gange. */
const OLD_LAUNCH_AGENT = join(homedir(), "Library", "LaunchAgents", "dk.miai.app.plist");

function migrateOldLaunchAgent() {
  if (!existsSync(OLD_LAUNCH_AGENT)) return;
  try {
    // Ingen "launchctl bootout": den ville stoppe appen selv, hvis den er startet af agenten.
    // Agenten er indlæst til næste log ud; uden filen indlæses den ikke igen.
    unlinkSync(OLD_LAUNCH_AGENT);
    app.setLoginItemSettings({ openAtLogin: true });
    log("Gammel LaunchAgent erstattet af login-element");
  } catch (err) {
    log(`Kunne ikke fjerne gammel LaunchAgent: ${err.message}`);
  }
}

function startAtLoginEnabled() {
  return app.isPackaged && app.getLoginItemSettings().openAtLogin;
}

/** Slår login-elementet til/fra. Kræver macOS godkendelse, vises der en vej til Indstillinger. */
async function setStartAtLogin(enabled) {
  // Fra en DMG eller Overførsler kører appen fra en midlertidig sti (App Translocation),
  // som et login-element ikke kan pege på – flyt først appen til Programmer.
  if (enabled && process.platform === "darwin" && !app.isInApplicationsFolder()) {
    const r = await ask({
      type: "question",
      message: "Flyt MIAI til mappen Programmer?",
      detail: "MIAI kan kun starte automatisk ved login, når appen ligger i Programmer. Appen genstarter derfra – slå så indstillingen til igen.",
      buttons: ["Flyt til Programmer", "Annullér"],
      defaultId: 0,
      cancelId: 1,
    });
    try {
      if (r.response === 0 && app.moveToApplicationsFolder()) return; // appen genstarter fra Programmer
    } catch (err) {
      log(`Flytning til Programmer fejlede: ${err.message}`);
      dialog.showErrorBox("MIAI", `MIAI kunne ikke flyttes til Programmer: ${err.message}. Træk appen derind i Finder.`);
    }
    refreshMenu();
    return;
  }
  app.setLoginItemSettings({ openAtLogin: enabled });
  const status = app.getLoginItemSettings().status;
  if (enabled && process.platform === "darwin" && status === "requires-approval") {
    const r = await ask({
      type: "info",
      message: "Godkend MIAI under Login-emner",
      detail: "macOS skal godkende, at MIAI starter ved login: Systemindstillinger → Generelt → Login-emner.",
      buttons: ["Åbn Indstillinger", "Luk"],
      defaultId: 0,
    });
    if (r.response === 0) void shell.openExternal("x-apple.systempreferences:com.apple.LoginItems-Settings.extension");
  }
  refreshMenu();
}

// --- En anden instans blev startet --------------------------------------------------
let secondInstanceOpen = false;

async function onSecondInstance(other) {
  if (secondInstanceOpen) return; // flere klik giver ikke flere dialoger
  const d = secondInstanceDialog(app.getVersion(), other, process.platform);
  log(`anden instans startet (${other?.version ?? "ukendt udgave"}): ${d.kind}`);
  secondInstanceOpen = true;
  try {
    const r = await ask({ type: "info", title: "MIAI", message: d.message, detail: d.detail, buttons: d.buttons, defaultId: 0 });
    if (d.kind === "switch" && r.response === 0) {
      log(`skifter til ${other.version}: ${d.execPath}`);
      app.relaunch({ execPath: d.execPath, args: [] });
      app.quit();
    }
  } finally {
    secondInstanceOpen = false;
  }
}

// --- Livscyklus -----------------------------------------------------------------
// Den nye instans fortæller den kørende, hvilken udgave den er, så den kørende kan sige det til
// brugeren i stedet for at lade den nye lukke uden et ord.
const gotLock = app.requestSingleInstanceLock(launchInfo(app.getVersion(), process.execPath));
if (!gotLock) {
  app.quit();
} else {
  app.on("second-instance", (_event, _argv, _cwd, other) => void onSecondInstance(other));
  // Appen lever i menulinjen; at et dialogvindue lukkes, må ikke afslutte den (Electrons standard).
  app.on("window-all-closed", () => {});

  process.on("unhandledRejection", (r) => log(`unhandledRejection: ${r}`));
  process.on("uncaughtException", (e) => log(`uncaughtException: ${e.stack ?? e}`));

  app
    .whenReady()
    .then(async () => {
      log("whenReady");
      // Kun menu-bar – ingen dock-ikon.
      app.dock?.hide();

      try {
        await startBridge();
      } catch (err) {
        log(`startBridge FEJL: ${err.stack ?? err}`);
        dialog.showErrorBox("MIAI", `Bridgen kunne ikke startes: ${err.message}`);
        app.quit();
        return;
      }

      const icon = nativeImage.createFromPath(join(app.getAppPath(), "assets", platform.trayIcon.file));
      if (platform.trayIcon.template) icon.setTemplateImage(true);
      tray = new Tray(icon.isEmpty() ? nativeImage.createEmpty() : icon);
      log("tray oprettet");
      if (app.isPackaged && process.platform === "darwin") migrateOldLaunchAgent();
      // Automatiske opdateringer (kun den installerede app; menuen viser, når en er klar).
      setupUpdates({ log, onChange: refreshMenu });
      // Udviklerudgaven: menupunkt, når der ligger en nyere på feedet (Mac på Apple Silicon og Windows x64).
      setupDevBuilds({ log, onChange: refreshMenu });

      // Opsætning kun når der mangler noget: certifikat (nyt eller fornyet) og/eller Office-installation.
      // Tidligere blev velkomstdialogen vist ved hver start.
      const needTrust = !isCertTrusted();
      const installState = officeInstallState();
      if (installState === "stale") {
        // Port eller adgangstoken er ændret: opdater manifestet i Office uden at spørge (det er brugerens egen installation).
        try {
          installInOffice();
          log("manifest opdateret i Office (ny port/token)");
          ask({
            type: "info",
            title: "MIAI",
            message: "MIAI er opdateret i Office.",
            detail: "Genstart Word, Excel og PowerPoint, så de bruger den nye forbindelse.",
          });
        } catch (err) {
          // Før blev det kun logget, og Office brugte i stilhed den gamle adresse, hvor intet svarer.
          log(`kunne ikke opdatere manifest: ${err.message}`);
          showInstallError(err, "MIAI har fået en ny adresse, men kunne ikke opdatere sig selv i Office. Indtil det er gjort, kan Office ikke nå MIAI.");
        }
      }
      const needInstall = installState === "missing";
      log(`opsætning: certifikat betroet=${!needTrust}, installeret i Office=${!needInstall}`);
      if (needTrust || needInstall) {
        const steps = [
          needTrust && "betro MIAI-certifikatet (bruges til den lokale HTTPS-forbindelse)",
          needInstall && "installere MIAI i Word, Excel og PowerPoint",
        ].filter(Boolean);
        const r = await ask({
          type: "info",
          title: "Velkommen til MIAI",
          message: `MIAI Bridge kører nu på https://localhost:${PORT}`,
          detail:
            `MIAI skal ${steps.join(" og ")}.` +
            (needTrust && process.platform === "win32" ? "\n\nWindows spørger, om du vil installere certifikatet – vælg Ja." : "") +
            "\n\nVælg derefter dit inference-værktøj og model under Indstillinger i add-in'et.",
          buttons: ["Gør det nu", "Senere"],
          defaultId: 0,
        });
        if (r.response === 0) {
          // Én dialog ad gangen.
          if (needTrust) await showCertResult(trustCert());
          if (needInstall) {
            try {
              installInOffice();
              ask({
                type: "info",
                title: "MIAI",
                message: "MIAI er installeret i Office.",
                detail: "Genstart Word, Excel og PowerPoint. Du finder MIAI som 'AI-assistent' på fanen Hjem.",
              });
            } catch (err) {
              showInstallError(err);
            }
          }
        }
      }
    })
    .catch((err) => {
      log(`whenReady FEJL: ${err.stack ?? err}`);
      dialog.showErrorBox("MIAI", `Uventet fejl ved start: ${err.message}`);
    });

  app.on("before-quit", () => {
    quitting = true;
    bridgeProc?.kill();
  });
}
