import { chunks, expect, test } from "./fixtures/test";

/** Chat: figurer, gentagelsessløjfer og indstillingen presence_penalty. */

test("Mermaid i svaret vises som figur, teksten bevares, og kan indsættes som billede", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  const answer =
    "Her er processen:\n\n```mermaid\nflowchart LR\n  A[Indsamling af data] --> B[Analyse]\n  B --> C[Rapport til ledelsen]\n```\n\nSig til, hvis den skal justeres.";
  tp.model.reply = () => ({ stream: chunks(answer) });

  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Tegn processen");
  await tp.page.keyboard.press("Enter");

  const figure = tp.page.locator(".figure svg").first();
  await expect(figure).toBeVisible({ timeout: 15_000 });
  // htmlLabels: false → ren SVG-tekst (WebKit kan ikke eksportere <foreignObject> til PNG).
  expect(await figure.locator("foreignObject").count()).toBe(0);
  await expect(figure).toContainText("Analyse");
  // Regression: figuren erstattede tidligere hele beskeden.
  await expect(tp.page.getByText("Her er processen:")).toBeVisible();
  await expect(tp.page.getByText("Sig til, hvis den skal justeres.")).toBeVisible();

  await tp.page.getByRole("button", { name: "Indsæt som billede" }).click();
  await expect(tp.page.getByText("Figuren er indsat efter markeringen.")).toBeVisible();
  const pic = (await tp.find<{ width: number; height: number }>("picture"))!;
  // Mermaids viewBox (negative tal) læses korrekt: bred figur, ikke 600x400-standarden.
  expect(pic.width / pic.height).toBeGreaterThan(3);
});

test("gentagelsessløjfe afbrydes med advarsel – uden note i selve teksten", async ({ open }) => {
  const tp = await open({ host: "Word" });
  tp.model.reply = () => ({ stream: ["Jeg kigger på indholdet. ", ...Array(10).fill("side 6 tabel ")] });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Lav en figur");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText(/modellen begyndte at gentage sig selv/)).toBeVisible();
  await expect(tp.page.locator(".fui-MessageBar").filter({ hasText: "gentage sig selv" })).toHaveCount(1);
});

test("presence_penalty vælges i Indstillinger og sendes kun til fritekst", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.page.getByRole("combobox").filter({ hasText: "Fra (serverens standard)" }).click();
  await tp.page.getByRole("option", { name: /Kraftig/ }).click();
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();

  tp.model.reply = () => ({ stream: ["Hej med dig."] });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hej");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Hej med dig.")).toBeVisible();
  expect(tp.model.lastRequest?.presence_penalty).toBe(1.5);
});

test("web-søgning: slået fra som standard i bridge-tilstand, værktøjet sendes ikke", async ({ open }) => {
  const tp = await open({ host: "Word", mode: "bridge" });
  tp.model.reply = () => ({ stream: ["Svar uden søgning."] });
  await tp.tab("Indstillinger");
  await expect(tp.page.getByRole("switch", { name: "Web-søgning" })).not.toBeChecked();
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad er nyt?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Svar uden søgning.")).toBeVisible();
  expect(tp.model.lastRequest).not.toHaveProperty("tools");
});
