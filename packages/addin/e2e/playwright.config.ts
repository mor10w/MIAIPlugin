import { defineConfig, devices } from "@playwright/test";

/**
 * Browsertests af task pane (#10): add-in'ets build køres i Chromium med simuleret Office.js
 * og en mock OpenAI-server. Kræver `npm run build` først (testene bruger dist/ og bridgens dist/).
 */
export default defineConfig({
  testDir: ".",
  testMatch: "**/*.spec.ts",
  outputDir: "../test-results",
  timeout: 30_000,
  expect: {
    timeout: 10_000,
    // Visuelle tests (#17): fast skrifttype, ingen animationer; små forskelle i kantudglatning tolereres.
    toHaveScreenshot: { animations: "disabled", caret: "hide", maxDiffPixelRatio: 0.01 },
  },
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? [["list"], ["html", { open: "never", outputFolder: "../playwright-report" }]] : "list",
  globalSetup: "./global-setup.ts",
  use: {
    ...devices["Desktop Chrome"],
    // Task pane-størrelse som i Office.
    viewport: { width: 380, height: 900 },
    ignoreHTTPSErrors: true,
    locale: "da-DK",
    screenshot: "only-on-failure",
    trace: "retain-on-failure",
  },
});
