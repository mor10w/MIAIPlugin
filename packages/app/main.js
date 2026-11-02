// MIAI App – menu-bar-app, der kører MIAI Bridge på https://localhost:44321.
// Salgsudgave B (ADR-0008): alt kører lokalt; ingen vendor-cloud, ingen CORS.
// Bridgen (Fastify) kører i en Electron utility-process – brugeren behøver ikke Node.js.
// VIGTIGT: bridgen kan IKKE køre i main-processen – Chromiums netværksspecifik forstyrrer
// Node's TLS-server, så håndskaber hænger. I en utility-process (uden Chromium) fungerer
// TLS som i almindelig Node – verificeret med ELECTRON_RUN_AS_NODE + curl.
import { app, dialog, Menu, Tray, nativeImage, shell, utilityProcess } from "electron";
import { spawnSync } from "node:child_process";
import { appendFileSync, copyFileSync, mkdirSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { VERSION as BRIDGE_VERSION } from "@miai/bridge/dist/server.js";

const PORT = Number(process.env.MIAI_PORT ?? 44321);
const MIAI_DIR = join(homedir(), "Library", "Application Support", "MIAI");
const CERT_FILE = join(MIAI_DIR, "tls", "localhost.crt");

/** Fillog: Electron svulger console-output i pakker, så fejl kan ellers ikke ses. */
function log(msg) {
  try {
    appendFileSync(join(MIAI_DIR, "MIAI.log"), `[${new Date().toISOString()}] ${msg}\n`);
  } catch { /* log skal aldrig fælde appen */ }
}

// Packaget: filer ligger i Resources/ (electron-builder extraResources). I dev: repoet.
const ADDIN_DIR = app.isPackaged ? join(process.resourcesPath, "addin") : new URL("../../addin/dist", import.meta.url).pathname;
const MANIFEST = app.isPackaged ? join(process.resourcesPath, "manifest.xml") : new URL("../../addin/manifest.xml", import.meta.url).pathname;

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
    { env: { ...process.env, MIAI_PORT: String(PORT), MIAI_ADDIN_DIR: ADDIN_DIR } },
  );

  bridgeProc.on("message", (msg) => {
    if (msg?.type === "listening") {
      bridgeState.listening = true;
      log(`bridge-process: lytter på ${msg.addr} (v${msg.version})`);
      setStatus("Bridgen kører");
    } else if (msg?.type === "status") {
      setStatus(statusFromEndpoint(msg.endpoint));
    } else if (msg?.type === "error") {
      log(`bridge-process FEJL: ${msg.message}`);
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
    bridgeProc.on("message", (m) => { if (m?.type === "listening") { clearTimeout(t); resolve(); } });
    bridgeProc.on("exit", () => { clearTimeout(t); resolve(); });
  });
}

// --- Office -----------------------------------------------------------------
const OFFICE_APPS = ["Word", "Excel", "Powerpoint"];

/** Sideloader bridge-manifestet i de tre Office-apps (som scripts/sideload-mac.mjs). */
function installInOffice() {
  const xml = readFileSync(MANIFEST, "utf8");
  const id = xml.match(/<Id>([^<]+)<\/Id>/)?.[1];
  if (!id) throw new Error("Fandt ikke <Id> i manifestet");
  for (const app_ of OFFICE_APPS) {
    const dir = join(homedir(), "Library", "Containers", `com.microsoft.${app_}`, "Data", "Documents", "wef");
    mkdirSync(dir, { recursive: true });
    copyFileSync(MANIFEST, join(dir, `miai-${id}.xml`));
  }
}

/**
 * Betror certifikatet i login-keychain. Kræver ikke admin-adgang (egen login-keychain),
 * så der vises ingen macOS-dialog. Returnerer { ok, message } med den reelle fejl ved fiasko.
 */
function trustCert() {
  const keychain = join(homedir(), "Library", "Keychains", "login.keychain-db");
  const r = spawnSync("security", ["add-trusted-cert", "-d", "-r", "trustRoot", "-k", keychain, CERT_FILE], { encoding: "utf8" });
  return { ok: r.status === 0, message: (r.stderr ?? "").trim() };
}

function showCertResult(r) {
  dialog.showMessageBox({
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
          dialog.showMessageBox({
            type: "info",
            title: "MIAI er installeret i Office",
            message: "Manifestet er kopieret til Word, Excel og PowerPoint.",
            detail: "Genstart de Office-apps, du vil bruge MIAI i. Siden da finder du MIAI under Indsæt → Tilføjelser (og som 'AI-assistent' på fanen Hjem).",
          });
        } catch (err) {
          dialog.showErrorBox("MIAI", `Kunne ikke installere i Office: ${err.message}`);
        }
      },
    },
    {
      label: "Betro certifikat…",
      click: () => showCertResult(trustCert()),
    },
    { label: `Åbn https://localhost:${PORT}`, click: () => void shell.openExternal(`https://localhost:${PORT}`) },
    { type: "separator" },
    { label: "Afslut MIAI", click: () => app.quit() },
  ]);
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

// --- Livscyklus -----------------------------------------------------------------
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on("second-instance", () => {});

  process.on("unhandledRejection", (r) => log(`unhandledRejection: ${r}`));
  process.on("uncaughtException", (e) => log(`uncaughtException: ${e.stack ?? e}`));

  app.whenReady()
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

      const icon = nativeImage.createFromPath(join(app.getAppPath(), "assets", "tray.png"));
      icon.setTemplateImage(true);
      tray = new Tray(icon.isEmpty() ? nativeImage.createEmpty() : icon);
      log("tray oprettet");

      // Første start: vejledning + ét-klik-betroning.
      dialog.showMessageBox({
        type: "info",
        title: "Velkommen til MIAI",
        message: "MIAI Bridge kører nu på https://localhost:44321",
        detail:
          "1) Klik 'Betro certifikat' i menuen (ét klik – ingen admin-adgang).\n" +
          "2) Klik 'Installer i Office' og genstart Word/Excel/PowerPoint.\n" +
          "3) Vælg dit inference-værktøj og model i add-in'ets fanen Indstillinger.",
        buttons: ["Betro certifikat", "Senere"],
      }).then((r) => { if (r.response === 0) showCertResult(trustCert()); });
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
