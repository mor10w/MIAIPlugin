// MIAI Bridge i en Electron utility-process (separat fra Chromium).
// Bridgen KAN IKKE køre i Electron's main-process: Chromiums netværksspecifik
// forstyrrer Node's TLS-server, så håndskaber hænger. I en utility-process
// (uden Chromium) fungerer TLS som i almindelig Node – verificeret med curl.
import { execFileSync } from "node:child_process";
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import https from "node:https";
import { homedir } from "node:os";
import { join } from "node:path";
import { ConfigStore } from "@miai/bridge/dist/config.js";
import { buildServer, VERSION as BRIDGE_VERSION } from "@miai/bridge/dist/server.js";

const PORT = Number(process.env.MIAI_PORT ?? 44321);
// parentPort eksisterer kun i en utility-process (ikke export fra 'electron').
const parentPort = process.parentPort;
const MIAI_DIR = join(homedir(), "Library", "Application Support", "MIAI");
const TLS_DIR = join(MIAI_DIR, "tls");
const CERT_FILE = join(TLS_DIR, "localhost.crt");

function log(msg) {
  try {
    appendFileSync(join(MIAI_DIR, "MIAI.log"), `[${new Date().toISOString()}] [bridge-process] ${msg}\n`);
  } catch { /* log skal aldrig fælde bridgen */ }
}

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

/** Henter /api/health med TLS-verifikation slået fra (eget selvsigneret cert). */
function health() {
  return new Promise((resolve) => {
    const req = https.get(
      { host: "127.0.0.1", port: PORT, path: "/api/health", headers: { "X-MIAI": "1" }, rejectUnauthorized: false, timeout: 5_000 },
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

/** Rapporterer model-status til main-processen (menuens statuslinje). */
async function reportStatus() {
  const h = await health();
  log(`status-polling: ${h ? JSON.stringify(h.endpoint ?? h) : "null"}`);
  parentPort?.postMessage({ type: "status", endpoint: h?.endpoint ?? null });
}

try {
  log("starter");
  const store = new ConfigStore();
  const tls = ensureTls();
  log(`TLS OK (cert ${tls.cert.length} bytes)`);
  const server = await buildServer({ store, port: PORT, addinDir: process.env.MIAI_ADDIN_DIR, https: tls });
  const addr = await server.listen({ host: "127.0.0.1", port: PORT });
  log(`lytter på ${addr}`);
  parentPort?.postMessage({ type: "listening", addr, version: BRIDGE_VERSION });
  await reportStatus();
  setInterval(() => void reportStatus(), 20_000);
} catch (err) {
  log(`FEJL: ${err?.stack ?? err}`);
  parentPort?.postMessage({ type: "error", message: String(err?.message ?? err) });
  process.exit(1);
}
