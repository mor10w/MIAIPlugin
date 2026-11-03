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
      apiSets: { WordApi: "1.6" },
      paragraphs: [
        { text: "Kære kunde", style: "Normal" },
        { text: "Tak for din henvendelse. Vores kunder er vigtige for os.", style: "Normal" },
        { text: "Gammel note: ring til Bo.", style: "Normal", bold: true },
        { text: "Med venlig hilsen", style: "Normal" },
        { text: "Kundeservice", style: "Normal" },
      ],
    });
    tp.model.reply = () => ({
      json: JSON.stringify({
        summary: "Retter 'kunder' til 'klienter', tilføjer en overskrift og fjerner den gamle note.",
        actions: [
          { action: "replace", find: "kunder", replace: "klienter", all: true, reason: "Det ønskede ordvalg" },
          { action: "insert", text: "# Svar på din henvendelse", position: "start", reason: "Overskrift øverst" },
          { action: "delete", paragraph: 3, reason: "Intern note, der ikke skal med" },
          { action: "setStyle", paragraph: 5, style: "quote", reason: "Fremhæv afsenderen" },
          { action: "comment", find: "Kundeservice", text: "Skal der stå et navn?", reason: "Forslag" },
        ],
      }),
    });
    await tp.run("Agent: redigér dokumentet");
    await expect(tp.page.getByRole("button", { name: "Udfør 5 handlinger" })).toBeVisible();
    await tp.page.getByRole("button", { name: "Udfør 5 handlinger" }).scrollIntoViewIfNeeded();
    await settle(tp.page);
    await tp.page.screenshot({ path: `${OUT}/7-agent-plan-${theme}.png`, animations: "disabled", fullPage: true });
    await tp.page.getByRole("button", { name: "Udfør 5 handlinger" }).click();
    await expect(tp.page.getByText("5 af 5 handlinger udført.")).toBeVisible();
    await tp.page.getByRole("button", { name: "Fortryd agent-session" }).scrollIntoViewIfNeeded();
    await settle(tp.page);
    await tp.page.screenshot({ path: `${OUT}/8-agent-udfoert-${theme}.png`, animations: "disabled", fullPage: true });
  });
}

for (const style of ["clean", "bold", "editorial"] as const) {
  test(`billeder: design – ${style}`, async ({ open }) => {
    const tp = await open({
      host: "PowerPoint",
      apiSets: { PowerPointApi: "1.8" },
      slides: [{ layout: "Title Slide", shapes: [{ name: "Title 1", ph: "Title", text: "Kvartalsrapport" }] }],
    });
    const palettes = {
      clean: { background: "#FFFFFF", text: "#1B2A3A", muted: "#5B6B7B", primary: "#0A5C8A", accent: "#E07A1F" },
      bold: { background: "#F7F7F5", text: "#1D1D1F", muted: "#5E5E63", primary: "#1F3A5F", accent: "#2BB3A3" },
      editorial: { background: "#FBF8F3", text: "#2B2622", muted: "#6B625A", primary: "#7A2E2E", accent: "#B5651D" },
    };
    tp.model.reply = () => ({
      json: JSON.stringify({
        name: { clean: "Nordisk klar", bold: "Markant blå", editorial: "Varm redaktionel" }[style],
        rationale: "Rolig og troværdig til bestyrelsen; farverne understøtter tallene.",
        style,
        palette: palettes[style],
        fonts: style === "editorial" ? { heading: "Georgia", body: "Aptos" } : { heading: "Aptos Display", body: "Aptos" },
        slides: [
          { layout: "title", title: "Kvartalsrapport Q3", subtitle: "Bestyrelsesmøde · november" },
          {
            layout: "kpi",
            title: "Resultater",
            kpis: [
              { value: "+12 %", label: "Omsætning" },
              { value: "38 %", label: "Bruttomargin" },
              { value: "127", label: "Medarbejdere" },
            ],
          },
          { layout: "section", title: "Næste skridt" },
          {
            layout: "twoColumn",
            title: "Norge og Sverige",
            left: { heading: "Norge", bullets: ["To nye butikker", "Vækst 18 %"] },
            right: { heading: "Sverige", bullets: ["Webshop i marts", "Ny betalingsløsning"] },
          },
          { layout: "quote", quote: "Det bedste kvartal i virksomhedens historie", attribution: "CFO" },
          { layout: "closing", title: "Tak", subtitle: "Spørgsmål?" },
        ],
      }),
    });
    await tp.run("Design: master og layout");
    await expect(tp.page.getByRole("img", { name: "6. Afslutning" })).toBeVisible();
    await tp.page.setViewportSize({ width: 450, height: 1300 });
    await tp.page
      .getByText("Farver", { exact: false })
      .first()
      .scrollIntoViewIfNeeded()
      .catch(() => {});
    await tp.page.locator("[role=img]").first().scrollIntoViewIfNeeded();
    await settle(tp.page);
    await tp.page.screenshot({ path: `${OUT}/9-design-${style}.png`, animations: "disabled" });
  });
}
