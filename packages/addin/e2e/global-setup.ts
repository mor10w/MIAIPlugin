import { existsSync } from "node:fs";
import { join } from "node:path";
import { createRequire } from "node:module";
import { DIST, testCert } from "./fixtures/servers";

/** Tjekker forudsætninger én gang og laver testcertifikatet, før workerne starter. */
export default function globalSetup(): void {
  if (!existsSync(join(DIST, "taskpane.html"))) {
    throw new Error(`Add-in'et er ikke bygget (${DIST}). Kør først: npm run build`);
  }
  const require = createRequire(import.meta.url);
  const bridgeDist = join(require.resolve("@miai/bridge/package.json"), "..", "dist", "server.js");
  if (!existsSync(bridgeDist)) throw new Error("Bridgen er ikke bygget. Kør først: npm run build");
  testCert(); // sætter MIAI_E2E_CERT_DIR, som workerne arver
}
