import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

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

export default defineConfig(({ command, mode }) => ({
  plugins: [react()],
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
