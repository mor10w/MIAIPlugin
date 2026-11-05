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

test("diagram med parenteser og <b> i etiketterne repareres og kan indsættes (regression fra Word-chatten)", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  const answer = [
    "Her er arkitekturen:",
    "",
    "```mermaid",
    "graph TD",
    "    A[<b>Dataarkitekt</b><br/>Hvem ejer data?]:::header",
    "    subgraph P2 [Platform & Integration]",
    "        R2[<b>Platformarkitekt</b><br/><i>(eksisterende arkitekturgruppe)</i>]:::role",
    "        T4[Dataplatform og lagring]",
    "    end",
    "    A --> P2",
    "    R2 --- T4",
    "    classDef header fill:#1f4e79,color:white",
    "    classDef role fill:#2e75b6,color:white",
    "```",
  ].join("\n");
  tp.model.reply = () => ({ stream: chunks(answer) });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Visualisér arkitekturen");
  await tp.page.keyboard.press("Enter");

  const figure = tp.page.locator(".figure svg").first();
  await expect(figure).toBeVisible({ timeout: 15_000 });
  await expect(figure).toContainText("(eksisterende arkitekturgruppe)");
  // <b> bliver til rigtig fed skrift, ikke synlige mærker.
  expect(await figure.textContent()).not.toContain("<b>");
  await expect(tp.page.getByText("Diagrammet kunne ikke tegnes")).toHaveCount(0);
  await tp.page.getByRole("button", { name: "Indsæt som billede" }).click();
  await expect(tp.page.getByText("Figuren er indsat efter markeringen.")).toBeVisible();
});

test("et diagram, der ikke kan tegnes, kan ikke indsættes, og modellen kan bede om at rette det", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  tp.model.reply = () => ({ stream: chunks("```mermaid\ngraph TD\n  A --> --> B\n```") });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Tegn det");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Diagrammet kunne ikke tegnes")).toBeVisible({ timeout: 15_000 });
  await expect(tp.page.getByRole("button", { name: "Indsæt som billede" })).toHaveCount(0);

  tp.model.reply = () => ({ stream: chunks('```mermaid\ngraph TD\n  A["Start"] --> B["Slut"]\n```') });
  await tp.page.getByRole("button", { name: "Bed modellen rette diagrammet" }).click();
  await expect(tp.page.locator(".figure svg").first()).toContainText("Slut", { timeout: 15_000 });
  const last = tp.model.lastRequest!.messages.at(-1)!.content;
  expect(last).toContain("kan ikke tegnes");
  expect(last).toMatch(/Parse error|Syntax error/i);
  // Det nye svar kan tegnes, så den gamle fejl står kun ved det første svar.
  await expect(tp.page.getByRole("button", { name: "Bed modellen rette diagrammet" })).toHaveCount(0);
  await expect(tp.page.getByRole("button", { name: "Indsæt som billede" })).toBeVisible();
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
  // presence_penalty ligger under endpointens "Avanceret".
  await tp.page.getByRole("button", { name: "Avanceret" }).click();
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
  await tp.section("Privatliv og politik");
  await expect(tp.page.getByRole("switch", { name: "Web-søgning" })).not.toBeChecked();
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad er nyt?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Svar uden søgning.")).toBeVisible();
  expect(tp.model.lastRequest).not.toHaveProperty("tools");
});
