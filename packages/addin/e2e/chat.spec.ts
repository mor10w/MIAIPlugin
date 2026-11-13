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

test("Excel: celleadresser i svaret kan klikkes og markerer området – ikke i kode og ikke ord, der ligner celler (#22)", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Region", "Omsætning"],
        ["Nord", 120],
        ["Syd", 340],
      ],
      sheets: [{ sheet: "Budget", values: [["Region", "Budget"]] }],
    },
  });
  tp.model.reply = () => ({
    stream: chunks("Summen af B2:B3 er **460**, og den største står i B3. Budgettet står i Budget!B2. I Q1 2026 brugte vi `=SUM(B2:B3)`."),
  });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad er den samlede omsætning?");
  await tp.page.keyboard.press("Enter");
  // Modellen bedes pege på celler med adresser.
  await expect(tp.page.getByText("Summen af")).toBeVisible();
  expect(tp.model.lastRequest!.messages[0].content).toContain("the user can click an address to jump to it");
  // Tre henvisninger: området, cellen inden for data og adressen med arknavn. Ikke "Q1" og ikke formlen i koden.
  await expect(tp.page.locator("button.docref")).toHaveText(["B2:B3", "B3", "Budget!B2"]);

  const selected = () =>
    tp.page.evaluate(() =>
      (window as unknown as { __miai: { calls: { kind: string; address?: string }[] } }).__miai.calls.filter((c) => c.kind === "select").map((c) => c.address),
    );
  await tp.page.getByRole("button", { name: "B2:B3" }).click();
  await expect.poll(selected).toEqual(["Ark1!B2:B3"]);
  // En adresse på et andet ark skifter til arket og markerer cellen.
  await tp.page.getByRole("button", { name: "Budget!B2" }).click();
  await expect.poll(selected).toEqual(["Ark1!B2:B3", "Budget!B2"]);
});

test("Word: en Markdown-tabel i chatsvaret indsættes som en rigtig tabel (#24)", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  tp.model.reply = () => ({ stream: chunks("Her er priserne:\n\n| Plan | Pris |\n|---|---|\n| Basis | 499 kr |\n| Plus | 799 kr |\n") });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad koster de to planer?");
  await tp.page.keyboard.press("Enter");
  // Tabellen vises som tabel i samtalen …
  await expect(tp.page.getByRole("cell", { name: "799 kr" })).toBeVisible();
  // … og indsættes som HTML-tabel, som Word gør til en tabel.
  await tp.page.getByRole("button", { name: "Indsæt efter" }).click();
  const html = (await tp.find<{ html: string }>("insertHtml"))!.html;
  expect(html).toContain("<table>");
  expect(html).toMatch(/<th>Plan<\/th>\s*<th>Pris<\/th>/);
  expect(html).toMatch(/<td>Plus<\/td>\s*<td>799 kr<\/td>/);
});

test("chatten viser den aktuelle markering og følger den, når brugeren markerer noget andet (#26)", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Region", "Omsætning"],
        ["Nord", 120],
        ["Syd", 340],
      ],
    },
  });
  await tp.tab("Chat");
  await expect(tp.page.getByText("Markering: Ark1!A1:B3")).toBeVisible();
  await tp.page.evaluate(() => (window as unknown as { __miai: { selectRange: (a: string) => void } }).__miai.selectRange("B2:B3"));
  await expect(tp.page.getByText("Markering: Ark1!B2:B3")).toBeVisible();
  // Det er den nye markering, modellen får.
  tp.model.reply = () => ({ stream: ["Summen er 460."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad er summen?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Summen er 460.")).toBeVisible();
  expect(tp.model.lastRequest!.messages[0].content).toContain("(Ark1!B2:B3)");
  // Slås konteksten fra, vises markeringen ikke.
  await tp.page.getByRole("switch", { name: "Brug markering/dokument som kontekst" }).click();
  await expect(tp.page.getByText("Markering: Ark1!B2:B3")).toHaveCount(0);
});

test("Word: markeringen vises som antal ord, eller at hele dokumentet bruges (#26)", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  await tp.tab("Chat");
  await expect(tp.page.getByText("Intet markeret – hele dokumentet bruges")).toBeVisible();
  await tp.page.evaluate(() => (window as unknown as { __miai: { selectText: (t: string) => void } }).__miai.selectText("Tre små ord"));
  await expect(tp.page.getByText("Markering: 3 ord")).toBeVisible();
});

test("samtalen huskes pr. dokument og ryddes med 'Ny samtale' (#26)", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, documentUrl: "https://firma.sharepoint.com/Rapport.docx" });
  tp.model.reply = () => ({ stream: ["Rapporten handler om **levering**."] });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad handler rapporten om?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Rapporten handler om")).toBeVisible();

  // Task pane åbnes igen på samme dokument: samtalen er der, og den går med som historik.
  await tp.page.reload();
  await tp.tab("Chat");
  await expect(tp.page.getByText("Hvad handler rapporten om?")).toBeVisible();
  await expect(tp.page.getByText("Rapporten handler om")).toBeVisible();
  tp.model.reply = () => ({ stream: ["Den er på 12 sider."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvor lang er den?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Den er på 12 sider.")).toBeVisible();
  expect(tp.model.lastRequest!.messages.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);

  // "Ny samtale" sletter den også fra lageret.
  await tp.page.getByRole("button", { name: "Ny samtale" }).click();
  await tp.page.reload();
  await tp.tab("Chat");
  await expect(tp.page.getByPlaceholder("Spørg om dit dokument…")).toBeVisible();
  await expect(tp.page.getByText("Hvad handler rapporten om?")).toHaveCount(0);
});

test("samtalen huskes ikke for dokumenter uden adresse, eller når det er slået fra (#26)", async ({ open }) => {
  const stored = (page: import("@playwright/test").Page) => page.evaluate(() => Object.keys(localStorage).filter((k) => k.startsWith("miai.chat.")));
  const ask = async (tp: Awaited<ReturnType<typeof open>>) => {
    tp.model.reply = () => ({ stream: ["Svar."] });
    await tp.tab("Chat");
    await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Spørgsmål?");
    await tp.page.keyboard.press("Enter");
    await expect(tp.page.getByText("Svar.")).toBeVisible();
  };
  const unsaved = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  await ask(unsaved);
  expect(await stored(unsaved.page)).toEqual([]);

  const off = await open({ host: "Word", apiSets: { WordApi: "1.4" }, documentUrl: "C:\\Dokumenter\\Notat.docx", settings: { chatHistory: false } });
  await ask(off);
  expect(await stored(off.page)).toEqual([]);
});

const PNG_1X1 = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64");
type Part = { type: string; text?: string; image_url?: { url: string } };

test("vedhæftninger (#27): et billede og en tekstfil går med til modellen og ryddes efter afsendelse", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: { sheet: "Ark1", values: [["Region", "Omsætning"]] } });
  await tp.tab("Chat");
  await tp.page.locator('input[type="file"]').setInputFiles([
    { name: "tabel.png", mimeType: "image/png", buffer: PNG_1X1 },
    { name: "noter.txt", mimeType: "text/plain", buffer: Buffer.from("Budgettet for 2026 er 900.") },
  ]);
  await expect(tp.page.getByText("tabel.png")).toBeVisible();
  await expect(tp.page.getByText("noter.txt")).toBeVisible();
  // En vedhæftning kan fjernes igen, før beskeden sendes.
  await tp.page.locator('input[type="file"]').setInputFiles([{ name: "fortrudt.txt", mimeType: "text/plain", buffer: Buffer.from("x") }]);
  await tp.page.getByRole("button", { name: "Fjern fortrudt.txt" }).click();
  await expect(tp.page.getByText("fortrudt.txt")).toHaveCount(0);

  tp.model.reply = () => ({ stream: ["Billedet viser en tabel med to kolonner."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad viser billedet?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Billedet viser en tabel med to kolonner.")).toBeVisible();

  const content = tp.model.lastRequest!.messages.at(-1)!.content as unknown as Part[];
  expect(content.map((p) => p.type)).toEqual(["text", "image_url"]);
  expect(content[0].text).toContain("Hvad viser billedet?");
  expect(content[0].text).toContain('<attachment name="noter.txt">\nBudgettet for 2026 er 900.\n</attachment>');
  expect(content[1].image_url!.url).toMatch(/^data:image\/png;base64,/);
  // Skrivefeltet er ryddet, og beskeden i samtalen viser, hvad der var vedhæftet.
  await expect(tp.page.getByRole("button", { name: /^Fjern / })).toHaveCount(0);
  await expect(tp.page.getByText("tabel.png")).toBeVisible();

  // Næste besked: tekstfilen følger med i historikken, billedet sendes ikke igen.
  tp.model.reply = () => ({ stream: ["Budgettet var 900."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad var budgettet?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Budgettet var 900.")).toBeVisible();
  const msgs = tp.model.lastRequest!.messages;
  expect(typeof msgs.at(-1)!.content).toBe("string");
  expect(msgs[1].content).toContain("[image attached earlier: tabel.png]");
  expect(msgs[1].content).toContain("Budgettet for 2026 er 900.");
});

test("vedhæftninger (#27): en model uden vision giver en forståelig fejl, og PDF kræver bridgen", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" } });
  await tp.tab("Chat");
  const input = tp.page.locator('input[type="file"]');
  await input.setInputFiles([{ name: "rapport.pdf", mimeType: "application/pdf", buffer: Buffer.from("%PDF-1.4") }]);
  await expect(tp.page.getByText("rapport.pdf: PDF og Office-filer kan kun vedhæftes med MIAI Bridge eller MIAI App.")).toBeVisible();
  await input.setInputFiles([{ name: "program.exe", mimeType: "application/octet-stream", buffer: Buffer.from("MZ") }]);
  await expect(tp.page.getByText(/program\.exe: filtypen kan ikke vedhæftes/)).toBeVisible();

  await input.setInputFiles([{ name: "skærm.png", mimeType: "image/png", buffer: PNG_1X1 }]);
  tp.model.reply = (b) => (Array.isArray(b.messages.at(-1)!.content) ? { status: 400, error: "image input is not supported" } : { stream: ["ok"] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad står der?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText(/Modellen kunne ikke bruge billedet – den understøtter måske ikke billeder/)).toBeVisible();
});

test("vedhæftninger (#27): et billede går også med, når chatten laver en plan for dokumentet", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: { sheet: "Ark1", values: [["Region", "Omsætning"]] } });
  await tp.tab("Chat");
  await tp.page.locator('input[type="file"]').setInputFiles([{ name: "tabel.png", mimeType: "image/png", buffer: PNG_1X1 }]);
  await expect(tp.page.getByText("tabel.png")).toBeVisible();
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Indsætter tallene fra billedet.",
      actions: [{ action: "setValues", address: "A2", values: [["Nord", 120]], reason: "Fra billedet" }],
    }),
  });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Indsæt tabellen fra billedet i arket");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Indsætter tallene fra billedet.")).toBeVisible();
  const req = tp.model.lastRequest!;
  expect(req.messages[0].content).toContain("You are in agent mode");
  expect((req.messages.at(-1)!.content as unknown as Part[]).map((p) => p.type)).toEqual(["text", "image_url"]);
});
