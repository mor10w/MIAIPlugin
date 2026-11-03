#!/usr/bin/env node
import "./quiet.js";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { ConfigStore, configDir } from "./config.js";
import { loadIdentity, writePersonalManifest } from "./identity.js";
import { buildServer, VERSION } from "./server.js";

const here = dirname(fileURLToPath(import.meta.url));
const addinDir = process.env.MIAI_ADDIN_DIR ?? resolve(here, "../../addin/dist");
/** Bridge-manifestet (https://localhost:44321), som brugerens personlige manifest laves ud fra. */
const baseManifest = process.env.MIAI_MANIFEST ?? resolve(here, "../../addin/manifest.xml");

/**
 * TLS-certifikat. Office kræver HTTPS for add-ins. I udvikling bruges certifikatet fra
 * `npx office-addin-dev-certs install` (~/.office-addin-dev-certs), som også gør CA'en betroet.
 * Installeren (fase 1) laver et eget lokalt CA og sætter MIAI_TLS_CERT / MIAI_TLS_KEY.
 */
function loadTls(): { cert: Buffer; key: Buffer } {
  const devDir = join(homedir(), ".office-addin-dev-certs");
  const certFile = process.env.MIAI_TLS_CERT ?? join(devDir, "localhost.crt");
  const keyFile = process.env.MIAI_TLS_KEY ?? join(devDir, "localhost.key");
  if (!existsSync(certFile) || !existsSync(keyFile)) {
    console.error(`[miai] TLS-certifikat mangler (${certFile}).\n       Kør: npm run certs`);
    process.exit(1);
  }
  return { cert: readFileSync(certFile), key: readFileSync(keyFile) };
}

const store = new ConfigStore();
// Token og port pr. bruger (#11). MIAI_PORT låser porten; ellers huskes en ledig port.
const { token, port } = await loadIdentity(configDir(), process.env.MIAI_PORT ? Number(process.env.MIAI_PORT) : undefined);
const manifest = writePersonalManifest(configDir(), baseManifest, { token, port });
const app = await buildServer({ store, port, addinDir, token, https: loadTls(), fastify: { logger: { level: process.env.MIAI_LOG ?? "info" } } });

try {
  // Kun loopback – bridgen må aldrig være tilgængelig fra netværket.
  await app.listen({ host: "127.0.0.1", port });
  const ep = store.activeEndpoint();
  console.log(`[miai] MIAI Bridge ${VERSION} på https://localhost:${port}`);
  console.log(`[miai] Konfiguration: ${store.file}`);
  console.log(`[miai] API-nøgler: ${store.keyStorage === "keyring" ? "OS'ets nøglering" : "config.json (ingen nøglering)"}`);
  console.log(`[miai] Aktiv endpoint: ${ep.name} (${ep.baseUrl}) · politik: ${store.policySource}`);
  if (manifest) console.log(`[miai] Dit manifest (med adgangstoken): ${manifest}`);
  else console.warn(`[miai] Fandt ikke ${baseManifest} – intet manifest skrevet. Sæt MIAI_MANIFEST.`);
} catch (err) {
  console.error(err);
  process.exit(1);
}

for (const sig of ["SIGINT", "SIGTERM"] as const) process.on(sig, () => void app.close().then(() => process.exit(0)));
