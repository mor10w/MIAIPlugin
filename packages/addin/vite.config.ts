import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { collectPackages, packageRootOf, renderNotices } from "../../scripts/third-party-notices.mjs";

/**
 * Add-in'et bygges til statiske filer, som kan:
 *  - hostes statisk (GitHub Pages o.l.) → direkte forbindelse til modellen, eller
 *  - serveres af MIAI Bridge på https://localhost:44321 → bridge-tilstand.
 * Tilstanden vælges ved runtime (se src/backend/index.ts), så det er samme build.
 *
 * `npm run dev:direct` kører Vite på https://localhost:3000 med Office-udviklingscertifikatet.
 */
function devHttps() {
  const dir = join(homedir(), ".office-addin-dev-certs");
  const cert = join(dir, "localhost.crt");
  const key = join(dir, "localhost.key");
  return existsSync(cert) && existsSync(key) ? { cert: readFileSync(cert), key: readFileSync(key) } : undefined;
}

/**
 * Lægger licenserne ved siden af det byggede add-in: MIAI's egen (LICENSE.txt) og tredjepartspakkernes
 * (THIRD-PARTY-NOTICES.txt). Listen laves ud fra de moduler, der faktisk er pakket med, så den følger
 * afhængighederne af sig selv.
 */
function licenseFiles(): Plugin {
  return {
    name: "miai-license-files",
    apply: "build",
    generateBundle() {
      const roots = new Set<string>();
      for (const id of this.getModuleIds()) {
        const root = packageRootOf(id);
        if (root) roots.add(root);
      }
      const packages = collectPackages(roots);
      if (packages.length === 0) this.error("Ingen tredjepartspakker fundet – licensfilen ville blive tom.");
      const intro = "MIAI's add-in (task pane) er pakket sammen med de tredjepartspakker, der er listet nedenfor.\n";
      this.emitFile({ type: "asset", fileName: "THIRD-PARTY-NOTICES.txt", source: renderNotices(packages, intro) });
      this.emitFile({ type: "asset", fileName: "LICENSE.txt", source: readFileSync(new URL("../../LICENSE", import.meta.url), "utf8") });
    },
  };
}

export default defineConfig(({ command, mode }) => ({
  plugins: [react(), licenseFiles()],
  base: "./",
  server: command === "serve" ? { port: 3000, strictPort: true, https: devHttps() } : undefined,
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: mode === "development",
    minify: mode !== "development",
    chunkSizeWarningLimit: 1500, // serveres lokalt/statisk, så bundlestørrelse er uvæsentlig
    rollupOptions: { input: { taskpane: "taskpane.html" } },
  },
}));
