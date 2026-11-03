import AxeBuilder from "@axe-core/playwright";
import { expect, test } from "./fixtures/test";
import { settle, VIEWS } from "./fixtures/views";

/**
 * Tilgængelighed (#17): axe-tjek (WCAG 2.1 A/AA) af hovedvisningerne i lyst og mørkt tema,
 * at temaet følger Office' temaskift, og at layoutet holder i et smalt panel.
 */

for (const theme of ["light", "dark"] as const) {
  for (const [name, setup] of Object.entries(VIEWS)) {
    test(`axe: ${name} – ${theme}`, async ({ open }) => {
      const tp = await open({ host: "Word", theme, apiSets: { WordApi: "1.4" }, selection: "Tekst, der skal forenkles." });
      await setup(tp);
      await settle(tp.page);
      const result = await new AxeBuilder({ page: tp.page })
        .withTags(["wcag2a", "wcag2aa", "wcag21a", "wcag21aa"])
        // Fluents fokushåndtering (Tabster) indsætter skjulte fokus-vagter, der kun flytter fokus videre.
        // De er aria-hidden og fokuserbare med vilje (Microsofts eget mønster), ikke et reelt problem.
        .exclude("[data-tabster-dummy]")
        .analyze();
      const serious = result.violations.filter((v) => v.impact === "serious" || v.impact === "critical");
      expect(serious.map((v) => `${v.id}: ${v.help} (${v.nodes.map((n) => n.target.join(" ")).join(", ")})`)).toEqual([]);
    });
  }
}

test("temaet følger Office, når brugeren skifter tema undervejs", async ({ open }) => {
  const tp = await open({ host: "Word", theme: "light" });
  const bg = () => tp.page.locator("main").evaluate((el) => getComputedStyle(el.parentElement!).backgroundColor);
  const light = await bg();
  await tp.page.evaluate(() => (window as unknown as { __miai: { setTheme: (t: string) => void } }).__miai.setTheme("dark"));
  await expect.poll(bg).not.toBe(light);
  await tp.page.evaluate(() => (window as unknown as { __miai: { setTheme: (t: string) => void } }).__miai.setTheme("light"));
  await expect.poll(bg).toBe(light);
});

for (const width of [320, 450]) {
  test(`ingen vandret scroll ved ${width} px`, async ({ open }) => {
    const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, selection: "Tekst, der skal forenkles." });
    await tp.page.setViewportSize({ width, height: 800 });
    for (const [name, setup] of Object.entries(VIEWS)) {
      if (name !== "handlinger") {
        // Tilbage til udgangspunktet mellem visningerne.
        await tp.tab("Handlinger");
        const back = tp.page.getByRole("button", { name: "Alle handlinger" });
        if (await back.isVisible()) await back.click();
      }
      await setup(tp);
      await settle(tp.page);
      const overflow = await tp.page.evaluate(() => {
        const main = document.querySelector("main")!;
        return { doc: document.documentElement.scrollWidth - document.documentElement.clientWidth, main: main.scrollWidth - main.clientWidth };
      });
      expect(overflow, `${name} ved ${width} px`).toEqual({ doc: 0, main: 0 });
    }
  });
}
