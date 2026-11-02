import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Add-in'et bygges til statiske filer, som MIAI Bridge serverer på https://localhost:44321 (same origin som API'et).
export default defineConfig(({ mode }) => ({
  plugins: [react()],
  base: "./",
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: mode === "development",
    minify: mode !== "development",
    chunkSizeWarningLimit: 1500, // serveres lokalt, så bundlestørrelse er uvæsentlig
    rollupOptions: { input: { taskpane: "taskpane.html" } },
  },
}));
