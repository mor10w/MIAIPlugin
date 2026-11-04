// MIAI Bridge i en Electron utility-process (separat fra Chromium).
// Bridgen KAN IKKE køre i Electron's main-process: Chromiums netværksspecifik
// forstyrrer Node's TLS-server, så håndskaber hænger. I en utility-process
// (uden Chromium) fungerer TLS som i almindelig Node – verificeret med curl.
import "@miai/bridge/dist/quiet.js";
import { X509Certificate } from "node:crypto";
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import selfsigned from "selfsigned";
import { ConfigStore, configDir } from "@miai/bridge/dist/config.js";
import { loadIdentity, writePersonalManifest } from "@miai/bridge/dist/identity.js";
import { buildServer, VERSION as BRIDGE_VERSION } from "@miai/bridge/dist/server.js";

/** Port og adgangstoken pr. bruger (#11) – sættes ved start. MIAI_PORT låser porten. */
let PORT;
let TOKEN;
// parentPort eksisterer kun i en utility-process (ikke export fra 'electron').
const parentPort = process.parentPort;
/** Samme mappe som bridgen (macOS: ~/Library/Application Support/MIAI, Windows: %APPDATA%\\MIAI). */
const MIAI_DIR = configDir();
const TLS_DIR = join(MIAI_DIR, "tls");
const CERT_FILE = join(TLS_DIR, "localhost.crt");

function log(msg) {
  try {
    appendFileSync(join(MIAI_DIR, "MIAI.log"), `[${new Date().toISOString()}] [bridge-process] ${msg}\n`);
  } catch {
    /* log skal aldrig fælde bridgen */
  }
}

/** macOS afviser TLS-servercertifikater med længere gyldighed end 825 dage (også lokalt betroede). */
const CERT_DAYS = 825;
/** Forny i god tid, før certifikatet udløber. */
const RENEW_BEFORE_DAYS = 30;

/** Skal certifikatet (gen)oprettes? Mangler, for lang gyldighed (ældre versioner brugte 3650 dage) eller snart udløbet. */
function certNeedsRenewal(keyFile) {
  if (!existsSync(CERT_FILE) || !existsSync(keyFile)) return true;
  try {
    const c = new X509Certificate(readFileSync(CERT_FILE));
    const start = Date.parse(c.validFrom);
    const end = Date.parse(c.validTo);
    const day = 86_400_000;
    return !(end > start) || (end - start) / day > CERT_DAYS || (end - Date.now()) / day < RENEW_BEFORE_DAYS;
  } catch {
    return true;
  }
}

/**
 * Genererer et selvsigneret certifikat for localhost (CN=localhost, SAN: DNS+IP) med ren
 * JavaScript (selfsigned), så det virker ens på macOS og Windows uden openssl.
 * Office kræver HTTPS; certifikatet skal derefter betroes (menuen). Et fornyet certifikat skal
 * betroes igen – main-processen opdager det (platform.isCertTrusted).
 */
async function ensureTls() {
  const keyFile = join(TLS_DIR, "localhost.key");
  if (certNeedsRenewal(keyFile)) {
    log("opretter nyt TLS-certifikat");
    mkdirSync(TLS_DIR, { recursive: true });
    const now = new Date();
    const pems = await selfsigned.generate([{ name: "commonName", value: "localhost" }], {
      keySize: 2048,
      algorithm: "sha256",
      notBeforeDate: now,
      notAfterDate: new Date(now.getTime() + CERT_DAYS * 86_400_000),
      extensions: [
        { name: "basicConstraints", cA: false },
        { name: "keyUsage", digitalSignature: true, keyEncipherment: true },
        { name: "extKeyUsage", serverAuth: true },
        {
          name: "subjectAltName",
          altNames: [
            { type: 2, value: "localhost" },
            { type: 7, ip: "127.0.0.1" },
          ],
        },
      ],
    });
    writeFileSync(keyFile, pems.private, { mode: 0o600 });
    writeFileSync(CERT_FILE, pems.cert);
  }
  return { cert: readFileSync(CERT_FILE), key: readFileSync(keyFile) };
}

let server;
let lastStatus = "";

/**
 * Rapporterer model-status til main-processen (menuens statuslinje). Bruger Fastifys
 * inject – samme kode som /api/health, men uden netværk, TLS eller selvsigneret-undtagelse.
 * Logger kun ændringer, så MIAI.log ikke vokser med en linje hvert 20. sekund.
 */
async function reportStatus() {
  let endpoint = null;
  try {
    const res = await server.inject({ method: "GET", url: "/api/health", headers: { host: `localhost:${PORT}`, "x-miai": "1", "x-miai-token": TOKEN } });
    endpoint = res.json().endpoint ?? null;
  } catch (err) {
    log(`status-fejl: ${err?.message ?? err}`);
  }
  const s = JSON.stringify(endpoint);
  if (s !== lastStatus) log(`status: ${s}`);
  lastStatus = s;
  parentPort?.postMessage({ type: "status", endpoint });
}

try {
  log("starter");
  const store = new ConfigStore();
  const tls = await ensureTls();
  log(`TLS OK (cert ${tls.cert.length} bytes)`);
  const id = await loadIdentity(configDir(), process.env.MIAI_PORT ? Number(process.env.MIAI_PORT) : undefined);
  PORT = id.port;
  TOKEN = id.token;
  // Brugerens personlige manifest med port og token – det er dét, der installeres i Office.
  const manifest = process.env.MIAI_MANIFEST ? writePersonalManifest(configDir(), process.env.MIAI_MANIFEST, id) : undefined;
  // Outlook installeres manuelt (Hent tilføjelsesprogrammer → Mine → Tilføj fra fil) med dette manifest.
  if (process.env.MIAI_OUTLOOK_MANIFEST) writePersonalManifest(configDir(), process.env.MIAI_OUTLOOK_MANIFEST, id, "manifest-outlook.xml");
  server = await buildServer({ store, port: PORT, token: TOKEN, addinDir: process.env.MIAI_ADDIN_DIR, https: tls });
  const addr = await server.listen({ host: "127.0.0.1", port: PORT });
  log(`lytter på ${addr}`);
  parentPort?.postMessage({ type: "listening", addr, version: BRIDGE_VERSION, port: PORT, manifest });
  await reportStatus();
  setInterval(() => void reportStatus(), 20_000);
} catch (err) {
  log(`FEJL: ${err?.stack ?? err}`);
  parentPort?.postMessage({ type: "error", message: String(err?.message ?? err) });
  process.exit(1);
}
