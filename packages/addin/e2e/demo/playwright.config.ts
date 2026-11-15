import { defineConfig, devices } from "@playwright/test";

/**
 * Optager demovideoen (ikke en test). Efter `npm run build`:
 *   npx playwright test -c packages/addin/e2e/demo/playwright.config.ts
 * Filen hedder *.video.ts, så de almindelige browsertests ikke samler den op.
 */
export default defineConfig({
  testDir: ".",
  testMatch: "**/*.video.ts",
  outputDir: "../../demo-results",
  timeout: 240_000,
  workers: 1,
  reporter: "list",
  globalSetup: "../global-setup.ts",
  use: {
    ...devices["Desktop Chrome"],
    // Optagelsen følger viewporten (pixeltætheden tæller ikke med), så scenen selv er 1920×1080.
    viewport: { width: 1920, height: 1080 },
    actionTimeout: 15_000,
    ignoreHTTPSErrors: true,
    locale: "da-DK",
    video: { mode: "on", size: { width: 1920, height: 1080 } },
    trace: "off",
  },
});
