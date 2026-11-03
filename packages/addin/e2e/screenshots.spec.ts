import { chunks, expect, test } from "./fixtures/test";
import { settle } from "./fixtures/views";

/**
 * Dokumentationsbilleder af brugerfladen (#17) – ikke en test. Kører kun med MIAI_SCREENSHOTS=<mappe>:
 *   MIAI_SCREENSHOTS=../../docs/ui/efter npx playwright test -c e2e/playwright.config.ts screenshots
 */
const OUT = process.env.MIAI_SCREENSHOTS;
test.skip(!OUT, "Sæt MIAI_SCREENSHOTS for at tage dokumentationsbilleder");

const ANSWER =
  "Projektet er **på sporet**: budgettet holder, og leverancen sker i Q1.\n\n- Risiko: leverandørens kapacitet i december\n- Næste skridt: godkendelse i styregruppen\n\n```mermaid\nflowchart LR\n  A[Analyse] --> B[Godkendelse] --> C[Levering]\n```";

for (const theme of ["light", "dark"] as const) {
  test(`billeder – ${theme}`, async ({ open }) => {
    const tp = await open({
      host: "Word",
      theme,
      apiSets: { WordApi: "1.4" },
      selection: "Den foreliggende tekst er af en overordentlig kompliceret beskaffenhed, og den bør forenkles.",
    });
    const shot = async (name: string) => {
      await settle(tp.page);
      await tp.page.screenshot({ path: `${OUT}/${name}-${theme}.png`, animations: "disabled" });
    };
    await expect(tp.page.getByText("Forbundet", { exact: true })).toBeVisible();
    await shot("1-handlinger");

    tp.model.reply = () => ({ stream: chunks("Teksten er kompliceret og bør gøres enklere.") });
    await tp.run("Forenkl");
    await expect(tp.page.getByText("Teksten er kompliceret og bør gøres enklere.")).toBeVisible();
    await shot("2-resultat");

    tp.model.reply = () => ({ stream: chunks(ANSWER, 12) });
    await tp.tab("Chat");
    await tp.page.getByRole("textbox").first().fill("Hvordan står det til med projektet?");
    await tp.page.keyboard.press("Enter");
    await expect(tp.page.getByText("Næste skridt: godkendelse i styregruppen")).toBeVisible();
    await tp.page
      .locator(".figure svg, [data-figure] svg")
      .first()
      .waitFor({ timeout: 15_000 })
      .catch(() => {});
    await shot("3-chat");

    await tp.tab("Indstillinger");
    await tp.page.waitForTimeout(500);
    await shot("4-indstillinger");

    await tp.tab("Egne opgaver");
    await tp.page.waitForTimeout(300);
    await shot("5-opgaver");
  });
}

test("billeder – uden forbindelse", async ({ open }) => {
  const tp = await open({ host: "Word", endpoint: { baseUrl: "http://127.0.0.1:9/v1" } });
  await tp.page.waitForTimeout(800);
  await settle(tp.page);
  await tp.page.screenshot({ path: `${OUT}/6-uden-forbindelse-light.png`, animations: "disabled" });
});

for (const theme of ["light", "dark"] as const) {
  test(`billeder: agent – ${theme}`, async ({ open }) => {
    const tp = await open({
      host: "Word",
      theme,
      apiSets: { WordApi: "1.4" },
      document: "Kære kunde\n\nTak for din henvendelse. Vores kunder er vigtige for os.\n\nGammel note: ring til Bo.\n\nMed venlig hilsen\nKundeservice",
    });
    tp.model.reply = () => ({
      json: JSON.stringify({
        summary: "Retter 'kunder' til 'klienter', tilføjer en overskrift og fjerner den gamle note.",
        actions: [
          { action: "replace", find: "kunder", replace: "klienter", all: true, reason: "Det ønskede ordvalg" },
          { action: "insert", text: "Svar på din henvendelse", position: "start", style: "heading1", reason: "Overskrift øverst" },
          { action: "delete", find: "Gammel note: ring til Bo.", reason: "Intern note, der ikke skal med" },
          { action: "comment", find: "Kundeservice", text: "Skal der stå et navn?", reason: "Forslag" },
        ],
      }),
    });
    await tp.run("Agent: redigér dokumentet");
    await expect(tp.page.getByRole("button", { name: "Udfør 4 handlinger" })).toBeVisible();
    await tp.page.getByRole("button", { name: "Udfør 4 handlinger" }).scrollIntoViewIfNeeded();
    await settle(tp.page);
    await tp.page.screenshot({ path: `${OUT}/7-agent-plan-${theme}.png`, animations: "disabled", fullPage: true });
    await tp.page.getByRole("button", { name: "Udfør 4 handlinger" }).click();
    await expect(tp.page.getByText("4 af 4 handlinger udført.")).toBeVisible();
    await tp.page.getByRole("button", { name: "Fortryd agent-session" }).scrollIntoViewIfNeeded();
    await settle(tp.page);
    await tp.page.screenshot({ path: `${OUT}/8-agent-udfoert-${theme}.png`, animations: "disabled", fullPage: true });
  });
}
