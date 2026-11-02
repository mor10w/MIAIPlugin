#!/usr/bin/env node
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { ConfigStore } from "./config.js";
import { buildServer, VERSION } from "./server.js";

const here = dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.MIAI_PORT ?? 44321);
const addinDir = process.env.MIAI_ADDIN_DIR ?? resolve(here, "../../addin/dist");

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
const app = await buildServer({ store, port, addinDir, https: loadTls(), fastify: { logger: { level: process.env.MIAI_LOG ?? "info" } } });

try {
  // Kun loopback – bridgen må aldrig være tilgængelig fra netværket.
  await app.listen({ host: "127.0.0.1", port });
  const ep = store.activeEndpoint();
  console.log(`[miai] MIAI Bridge ${VERSION} på https://localhost:${port}`);
  console.log(`[miai] Konfiguration: ${store.file}`);
  console.log(`[miai] Aktiv endpoint: ${ep.name} (${ep.baseUrl}) · politik: ${store.policySource}`);
} catch (err) {
  console.error(err);
  process.exit(1);
}

for (const sig of ["SIGINT", "SIGTERM"] as const) process.on(sig, () => void app.close().then(() => process.exit(0)));
