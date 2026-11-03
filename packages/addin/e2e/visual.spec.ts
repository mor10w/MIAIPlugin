import { expect, test } from "./fixtures/test";
import { settle, useTestFont, VIEWS } from "./fixtures/views";

/**
 * Visuelle regressionstests (#17): hovedvisningerne i lyst og mørkt tema sammenlignes med
 * godkendte skærmbilleder i e2e/visual.spec.ts-snapshots/. Opdatér bevidst efter en designændring:
 *   npx playwright test -c e2e/playwright.config.ts visual --update-snapshots
 */
const SELECTION = "Den foreliggende tekst er af en overordentlig kompliceret beskaffenhed, og den bør forenkles.";

for (const theme of ["light", "dark"] as const) {
  for (const [name, setup] of Object.entries(VIEWS)) {
    test(`${name} – ${theme}`, async ({ open }) => {
      const tp = await open({ host: "Word", theme, apiSets: { WordApi: "1.4" }, selection: SELECTION });
      await setup(tp);
      await useTestFont(tp.page);
      await settle(tp.page);
      await expect(tp.page).toHaveScreenshot(`${name}-${theme}.png`, {
        // Mock-modellens port skifter fra kørsel til kørsel.
        mask: [tp.page.getByLabel("Basis-URL (inkl. /v1)")],
      });
    });
  }
}

test("uden forbindelse – light", async ({ open }) => {
  const tp = await open({ host: "Word", endpoint: { baseUrl: "http://127.0.0.1:9/v1" } });
  await expect(tp.page.getByRole("button", { name: "Kom i gang" })).toBeVisible();
  await useTestFont(tp.page);
  await settle(tp.page);
  await expect(tp.page).toHaveScreenshot("uden-forbindelse-light.png");
});
