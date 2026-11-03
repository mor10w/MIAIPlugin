import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { join } from "node:path";
import type { Page } from "@playwright/test";
import { chunks, expect, type Taskpane } from "./test";

/**
 * Fælles opsætning af hovedvisningerne til visuelle tests og tilgængelighedstjek (#17).
 * Hver visning bringer task pane i en bestemt, deterministisk tilstand.
 */

const require = createRequire(import.meta.url);
const fontDir = join(require.resolve("@fontsource/inter/package.json"), "..", "files");

/**
 * Fast skrifttype (Inter, bundlet som devDependency) i stedet for systemets – så skærmbilleder
 * ser ens ud på udviklermaskiner og i CI. Kun til visuelle tests.
 */
export async function useTestFont(page: Page): Promise<void> {
  await page.route("**/__e2e-font/*.woff2", (route) =>
    route.fulfill({ contentType: "font/woff2", body: readFileSync(join(fontDir, new URL(route.request().url()).pathname.split("/").pop()!)) }),
  );
  await page.addStyleTag({
    content: [400, 600]
      .map((w) => `@font-face{font-family:"MIAI Test";font-weight:${w};src:url(/__e2e-font/inter-latin-${w}-normal.woff2) format("woff2");}`)
      .join("")
      .concat(`*{font-family:"MIAI Test",sans-serif !important}`),
  });
  await page.evaluate(() => document.fonts.ready);
}

/** Musen væk fra fanerne (tooltips) og fokus væk fra felter (blinkende markør). */
export async function settle(page: Page): Promise<void> {
  await page.mouse.move(1, 1);
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  await page.waitForTimeout(250); // indtoning (160 ms) færdig
}

const ANSWER =
  "Projektet er **på sporet**: budgettet holder, og leverancen sker i Q1.\n\n- Risiko: leverandørens kapacitet i december\n- Næste skridt: godkendelse i styregruppen";

export const VIEWS: Record<string, (tp: Taskpane) => Promise<void>> = {
  handlinger: async (tp) => {
    await expect(tp.page.getByRole("button", { name: "Forenkl", exact: true })).toBeVisible();
  },
  resultat: async (tp) => {
    tp.model.reply = () => ({ stream: chunks("Teksten er kompliceret og bør gøres enklere.") });
    await tp.run("Forenkl");
    await expect(tp.page.getByRole("button", { name: "Erstat" })).toBeVisible();
  },
  chat: async (tp) => {
    tp.model.reply = () => ({ stream: chunks(ANSWER, 12) });
    await tp.tab("Chat");
    await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvordan står det til med projektet?");
    await tp.page.keyboard.press("Enter");
    await expect(tp.page.getByRole("button", { name: "Indsæt efter" })).toBeVisible();
  },
  "chat-tom": async (tp) => {
    await tp.tab("Chat");
    await expect(tp.page.getByText("Spørg om dit dokument", { exact: true })).toBeVisible();
  },
  indstillinger: async (tp) => {
    await tp.tab("Indstillinger");
    await expect(tp.page.getByLabel("Basis-URL (inkl. /v1)")).toBeVisible();
  },
  "egne-opgaver": async (tp) => {
    await tp.tab("Egne opgaver");
    await expect(tp.page.getByText("Ingen egne opgaver endnu")).toBeVisible();
  },
};
