// MIAI App – menu-bar-app, der kører MIAI Bridge på https://localhost:44321.
// Salgsudgave B (ADR-0008): alt kører lokalt; ingen vendor-cloud, ingen CORS.
// Bridgen (Fastify) kører i-process – brugeren behøver ikke Node.js.
import { app, dialog, Menu, Tray, nativeImage, shell } from "electron";
import { execFileSync, spawnSync } from "node:child_process";
import https from "node:https";
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { ConfigStore } from "@miai/bridge/dist/config.js";
import { buildServer, VERSION as BRIDGE_VERSION } from "@miai/bridge/dist/server.js";

const PORT = Number(process.env.MIAI_PORT ?? 44321);
const MIAI_DIR = join(homedir(), "Library", "Application Support", "MIAI");
const TLS_DIR = join(MIAI_DIR, "tls");
const CERT_FILE = join(TLS_DIR, "localhost.crt");

// Packaget: filer ligger i Resources/ (electron-builder extraResources). I dev: repoet.
const ADDIN_DIR = app.isPackaged ? join(process.resourcesPath, "addin") : new URL("../../addin/dist", import.meta.url).pathname;
const MANIFEST = app.isPackaged ? join(process.resourcesPath, "manifest.xml") : new URL("../../addin/manifest.xml", import.meta.url).pathname;

let server = null;
let tray = null;

// --- TLS -----------------------------------------------------------------
/**
 * Genererer et selvsigneret certifikat for localhost (CN=localhost, SAN: DNS+IP).
 * Office kræver HTTPS; certifikatet skal derefter betroes i keychain (menuen).
 */
function ensureTls() {
  const keyFile = join(TLS_DIR, "localhost.key");
  if (!existsSync(CERT_FILE) || !existsSync(keyFile)) {
    mkdirSync(TLS_DIR, { recursive: true });
    const cnf = join(TLS_DIR, "openssl.cnf");
    writeFileSync(
      cnf,
      [
        "[req]",
        "distinguished_name = dn",
        "x509_extensions = v3_req",
        "prompt = no",
        "[dn]",
        "CN = localhost",
        "[v3_req]",
        "subjectAltName = DNS:localhost,IP:127.0.0.1",
        "keyUsage = digitalSignature, keyEncipherment",
        "extendedKeyUsage = serverAuth",
      ].join("\n"),
    );
    // /usr/bin/openssl (LibreSSL) medleveres på macOS.
    execFileSync("/usr/bin/openssl", [
      "req", "-x509", "-newkey", "rsa:2048", "-nodes",
      "-days", "3650",
      "-keyout", keyFile, "-out", CERT_FILE, "-config", cnf,
    ], { stdio: "ignore" });
  }
  return { cert: readFileSync(CERT_FILE), key: readFileSync(keyFile) };
}

// --- Bridgen ---------------------------------------------------------------
async function startBridge() {
  const store = new ConfigStore();
  server = await buildServer({ store, port: PORT, addinDir: ADDIN_DIR, https: ensureTls(), fastify: { logger: false } });
  await server.listen({ host: "127.0.0.1", port: PORT });
}

/** Henter /api/health med TLS-verifikation slået fra (eget selvsigneret cert). */
function health() {
  return new Promise((resolve) => {
    const req = https.get(
      { host: "localhost", port: PORT, path: "/api/health", rejectUnauthorized: false, timeout: 4_000 },
      (res) => {
        let body = "";
        res.on("data", (c) => (body += c));
        res.on("end", () => { try { resolve(JSON.parse(body)); } catch { resolve(null); } });
      },
    );
    req.on("error", () => resolve(null));
    req.on("timeout", () => { req.destroy(); resolve(null); });
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
 * Betror certifikatet i login-keychain (kræver admin-adgang, vises af macOS).
 * Bemærk: stier med " kan ikke håndteres – usandsynligt på macOS.
 */
function trustCert() {
  const keychain = join(homedir(), "Library", "Keychains", "login.keychain-db");
  const cmd = `security add-trusted-cert -d -r trustRoot -k "${keychain}" "${CERT_FILE}"`;
  const r = spawnSync("osascript", ["-e", `do shell script ${JSON.stringify(cmd)} with administrator privileges`], { encoding: "utf8" });
  return r.status === 0;
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
      click: () => {
        const ok = trustCert();
        dialog.showMessageBox({
          type: ok ? "info" : "warning",
          title: "Certifikat",
          message: ok ? "MIAI-certifikatet er nu betroet." : "Certifikatet blev ikke betroet (afbrudt?).",
          detail: ok ? "Office kan nu forbinde til MIAI Bridge." : "Prøv igen, eller genstart appen.",
        });
      },
    },
    { label: `Åbn https://localhost:${PORT}`, click: () => void shell.openExternal(`https://localhost:${PORT}`) },
    { type: "separator" },
    { label: "Afslut MIAI", click: () => app.quit() },
  ]);
}

async function refreshStatus() {
  if (!tray) return;
  const h = await health();
  if (!h?.endpoint) {
    statusText = "Bridgen svarer ikke";
  } else if (h.endpoint.reachable) {
    statusText = `Forbundet · ${h.endpoint.baseUrl}`;
  } else {
    statusText = `Modellen svarer ikke (${h.endpoint.error ?? "ukendt fejl"})`;
  }
  tray.setToolTip(statusText);
  tray.setContextMenu(buildMenu());
}

// --- Livscyklus -----------------------------------------------------------------
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on("second-instance", () => {});

  app.whenReady().then(async () => {
    // Kun menu-bar – ingen dock-ikon.
    app.dock?.hide();

    try {
      await startBridge();
    } catch (err) {
      dialog.showErrorBox("MIAI", `Bridgen kunne ikke startes: ${err.message}`);
      app.quit();
      return;
    }

    const icon = nativeImage.createFromPath(join(app.getAppPath(), "assets", "tray.png"));
    icon.setTemplateImage(true);
    tray = new Tray(icon.isEmpty() ? nativeImage.createEmpty() : icon);

    await refreshStatus();
    setInterval(() => void refreshStatus(), 15_000);

    // Første start: vejledning + ét-klik-betroning.
    dialog.showMessageBox({
      type: "info",
      title: "Velkommen til MIAI",
      message: "MIAI Bridge kører nu på https://localhost:44321",
      detail:
        "1) Klik 'Betro certifikat' i menuen (kræver en gang admin-adgang).\n" +
        "2) Klik 'Installer i Office' og genstart Word/Excel/PowerPoint.\n" +
        "3) Vælg dit inference-værktøj og model i add-in'ets fanen Indstillinger.",
      buttons: ["Betro certifikat", "Senere"],
    }).then((r) => { if (r.response === 0 && trustCert()) void refreshStatus(); });
  });

  app.on("before-quit", async () => {
    try { await server?.close(); } catch { /* afslut alligevel */ }
  });
}
