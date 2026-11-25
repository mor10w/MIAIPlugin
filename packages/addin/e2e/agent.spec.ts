import { chunks, expect, test } from "./fixtures/test";

/** Agent-tilstand (#6): plan → bekræft → udfør, med fortryd af hele sessionen. */

const DOC = "Vi har mange kunder i Jylland. Alle kunder får besked.\nKontakt os på info@firma.dk.\nGammel note, der skal væk.";
const EDITS = (kinds: string[]) => (c: { kind: string }) => kinds.includes(c.kind);
const DOC_EDITS = EDITS(["replace", "insertText", "insertParagraph", "delete", "comment", "style", "insertOoxml"]);

/** Et lille dokument med overskrift, fed tekst, punkt, tabel og direkte formatering. */
const PARAS = [
  { text: "Kundeinformation", style: "Heading1" },
  { text: "Pris: 499 kr pr. md. for alle kunder.", style: "Normal", marks: [{ start: 0, end: 5, bold: true }] },
  { text: "Levering 3-5 dage", style: "ListBullet", list: true },
  { text: "Pris", style: "Normal", table: true },
  { text: "Gammel note, der skal væk.", style: "Normal", bold: true },
  { text: "Kontakt os på info@firma.dk.", style: "Normal" },
];
type DocP = {
  text: string;
  style: string;
  list: boolean;
  table: boolean;
  bold: boolean;
  marks: { start: number; end: number; bold?: boolean; italic?: boolean; link?: string }[];
};
const doc = (tp: { page: import("@playwright/test").Page }) => tp.page.evaluate(() => (window as unknown as { __miai: { doc: () => DocP[] } }).__miai.doc());
const ORIGINAL: DocP[] = PARAS.map((p) => ({ list: false, table: false, bold: false, marks: [], ...p }) as DocP);

const PLAN = {
  summary: "Opdaterer prisen, tilføjer leveringstekst og et punkt, rydder op og flytter kontaktlinjen.",
  actions: [
    { action: "replace", find: "Pris: 499 kr pr. md.", replace: "Pris: 599 kr pr. md.", reason: "Ny pris" },
    { action: "insert", text: "Vi leverer **hurtigt** – se [vilkår](https://firma.dk/vilkaar).", position: "after", paragraph: 1, reason: "Intro" },
    { action: "insert", text: "- Returnering inden 30 dage", position: "after", paragraph: 3, reason: "Nyt punkt" },
    { action: "setStyle", paragraph: 6, style: "heading2", reason: "Fremhæv kontakt" },
    { action: "delete", paragraph: 5, reason: "Forældet" },
    { action: "move", paragraph: 6, to: 2, position: "before", reason: "Kontakt først" },
    { action: "delete", paragraph: 4, reason: "Tabellen er overflødig" },
    { action: "comment", find: "info@firma.dk", text: "Stadig gyldig?", reason: "Tjek" },
  ],
};

test("Word: formateringen bevares – kun det ændrede ændres, ny tekst arver ikke naboens formatering, og fortryd afviser sessionens ændringer", async ({
  open,
}) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  tp.model.reply = () => ({ json: JSON.stringify(PLAN) });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();

  // Modellen ser afsnittene med typografi, liste og tabel.
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain(
    "¶1 [Heading1] Kundeinformation\n¶2 [Normal] Pris: 499 kr pr. md. for alle kunder.\n¶3 [ListBullet] (list) Levering 3-5 dage\n¶4 [Normal] (table) Pris",
  );

  // Planen: læsbar, strukturændring i tabellen fravalgt, intet ændret endnu.
  await expect(tp.page.getByRole("checkbox", { name: "Gør ¶6 “Kontakt os på info@firma.dk.” til overskrift 2" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Slet afsnittet ¶4 “Pris”" })).toBeDisabled();
  await expect(tp.page.getByText("Ville ændre en tabels opbygning")).toBeVisible();
  expect(await doc(tp)).toEqual(ORIGINAL);

  await tp.page.getByRole("button", { name: "Udfør 7 handlinger" }).click();
  await expect(tp.page.getByText("7 af 7 handlinger udført.")).toBeVisible();

  const after = await doc(tp);
  expect(after.map((p) => [p.style, p.text])).toEqual([
    ["Heading1", "Kundeinformation"],
    ["Normal", "Vi leverer hurtigt – se vilkår."],
    ["Heading2", "Kontakt os på info@firma.dk."],
    ["Normal", "Pris: 599 kr pr. md. for alle kunder."],
    ["ListBullet", "Levering 3-5 dage"],
    ["ListBullet", "Returnering inden 30 dage"],
    ["Normal", "Pris"],
  ]);
  // Erstat: kun "499" er skiftet ud – "Pris:" er stadig fed, og fed har ikke bredt sig.
  expect(after[3].marks).toEqual([{ start: 0, end: 5, bold: true }]);
  // Ny tekst efter en overskrift er brødtekst uden arvet fed – med egen fed og link.
  expect(after[1]).toMatchObject({ bold: false, list: false });
  expect(after[1].marks).toEqual([
    { start: 11, end: 18, bold: true },
    { start: 24, end: 30, link: "https://firma.dk/vilkaar" },
  ]);
  // Punktet er et punkt; tabellen er urørt; den slettede note efterlader intet tomt afsnit.
  expect(after[5]).toMatchObject({ list: true });
  expect(after[6]).toMatchObject({ table: true });
  expect(after.some((p) => p.text === "" || p.text.includes("Gammel note"))).toBe(false);

  // Fortryd: sessionens sporede ændringer afvises, og agentens kommentar fjernes.
  await expect(tp.page.getByText(/afviser agentens sporede ændringer/)).toBeVisible();
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await doc(tp)).toEqual(ORIGINAL);
  expect(await tp.page.evaluate(() => (window as unknown as { __miai: { comments: () => unknown[] } }).__miai.comments())).toEqual([]);
  // Øjebliksbilledet (hele dokumentet erstattet) er ikke brugt.
  expect((await tp.calls()).some((c) => c.kind === "insertOoxml" && c.loc === "Replace")).toBe(false);
});

test("Word uden WordApi 1.6: fortryd gendanner fra øjebliksbillede", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, paragraphs: PARAS });
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "", actions: [PLAN.actions[0], PLAN.actions[4]] }) });
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("2 af 2 handlinger udført.")).toBeVisible();
  await expect(tp.page.getByText(/også ændringer, du selv har lavet siden/)).toBeVisible();
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await doc(tp)).toEqual(ORIGINAL);
});

test("Word: ændres dokumentet mellem plan og udførelse, røres de berørte afsnit ikke", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "", actions: [{ action: "setStyle", paragraph: 6, style: "heading2", reason: "" }] }) });
  await tp.run("Agent: redigér dokumentet");
  await tp.page.evaluate(() =>
    (window as unknown as { __miai: { editParagraph: (i: number, t: string) => void } }).__miai.editParagraph(5, "Ring på 12 34 56 78."),
  );
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("Dokumentet er ændret siden planen – kør agenten igen")).toBeVisible();
  expect((await doc(tp))[5].style).toBe("Normal");
});

test("Excel (bridge): værdier og formel skrives på arket og fortrydes celle for celle", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    mode: "bridge",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Navn", "Beløb"],
        ["Anna", 1200],
        ["Bo", 800],
      ],
    },
  });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Tilføjer en total under beløbene.",
      actions: [
        { action: "setValues", address: "A4", values: [["Total"]], reason: "Etiket" },
        { action: "setFormula", address: "B4", formula: "=SUM(B2:B3)", reason: "Summen" },
        { action: "replace", find: "x", replace: "y", reason: "Forkert app" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Findes ikke i denne app")).toBeVisible();
  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("2 af 2 handlinger udført.")).toBeVisible();

  const writes = () => tp.calls().then((cs) => cs.filter((c) => c.kind === "values" || c.kind === "formulas"));
  expect(await writes()).toEqual([
    { kind: "values", address: "Ark1!A4", values: [["Total"]] },
    { kind: "formulas", address: "Ark1!B4", formulas: [["=SUM(B2:B3)"]] },
  ]);

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  // De tomme celler gendannes i omvendt rækkefølge.
  expect((await writes()).slice(2)).toEqual([
    { kind: "formulas", address: "Ark1!B4", formulas: [[""]] },
    { kind: "formulas", address: "Ark1!A4", formulas: [[""]] },
  ]);
});

test("ukendte værktøjer (fx netværk) afvises af skemaet – intet udføres", async ({ open }) => {
  const tp = await open({ host: "Word", document: DOC });
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "", actions: [{ action: "fetchUrl", url: "https://evil.example", reason: "" }] }) });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText(/gyldigt struktureret svar/)).toBeVisible();
  expect((await tp.calls()).filter(DOC_EDITS)).toEqual([]);
  await expect(tp.page.getByRole("button", { name: /Udfør/ })).toHaveCount(0);
});

test("English UI: the plan and confirmation are localized", async ({ open }) => {
  const tp = await open({ host: "Word", document: DOC });
  await tp.tab("Indstillinger");
  await tp.section("Generelt");
  await tp.page.getByRole("combobox", { name: "Sprog" }).click();
  await tp.page.getByRole("option", { name: "English" }).click();
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await tp.page.getByRole("tab", { name: "Actions" }).click();

  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Fix wording.", actions: [{ action: "delete", find: "Gammel note", reason: "Old" }] }) });
  await tp.page.getByRole("button", { name: "Agent: edit the document" }).click();
  await tp.page.getByRole("textbox").fill("Remove the old note");
  await tp.page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(tp.page.getByText("Proposed actions")).toBeVisible();
  await expect(tp.page.getByRole("checkbox", { name: "Delete “Gammel note”" })).toBeChecked();
  await tp.page.getByRole("button", { name: "Run 1 action" }).click();
  await expect(tp.page.getByText("1 of 1 actions done.")).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Undo agent session" })).toBeVisible();
  // The request asked for English reasons.
  expect(tp.model.lastRequest!.messages[0].content).toContain("Write summary and reasons in English");
});

test("chat → agent: forslaget vises ved ændringsønsker, og chatsvaret følger med som indhold", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  await tp.tab("Chat");
  // Et spørgsmål giver intet forslag.
  tp.model.reply = () => ({ stream: ["Vi mangler et afsnit om levering."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad mangler vi i dokumentet?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Vi mangler et afsnit om levering.")).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Gør det med agenten" })).toHaveCount(0);

  // Med "Må ændre dokumentet" slået fra svarer chatten med tekst, og et ændringsønske
  // (med æ i starten – regression) giver forslaget om at bruge agenten.
  await tp.page.getByRole("switch", { name: "Må ændre dokumentet" }).click();
  tp.model.reply = () => ({ stream: ["## Levering\nVi leverer på **3 dage**."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Ændr dokumentet, så det har et afsnit om levering");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Indsæt i dokumentet")).toBeVisible();
  await tp.page.getByRole("button", { name: "Gør det med agenten" }).click();

  // Agent-opgaven er udfyldt med ønsket og har chatsvaret med som indhold.
  await expect(tp.page.getByRole("heading", { name: "Agent: redigér dokumentet" })).toBeVisible();
  await expect(tp.page.getByRole("textbox")).toHaveValue("Ændr dokumentet, så det har et afsnit om levering");
  await expect(tp.page.getByText("Bruger chatsvaret som indhold.")).toBeVisible();
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Tilføjer afsnittet.", actions: [] }) });
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await expect(tp.page.getByText("Tilføjer afsnittet.")).toBeVisible();
  expect(tp.model.lastRequest!.messages[1].content).toContain("<material>\n## Levering\nVi leverer på **3 dage**.\n</material>");
});

test("chatten ændrer selv dokumentet: plan i samtalen, udførelse, og opfølgningen kender udfaldet (#19)", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  await tp.tab("Chat");
  const first = PARAS[1].text.split(" ").slice(0, 2).join(" ");
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Retter ordlyden i første afsnit.",
      actions: [{ action: "replace", find: first, replace: "Ny begyndelse", reason: "Ønsket" }],
    }),
  });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Ændr begyndelsen af første afsnit");
  await tp.page.keyboard.press("Enter");

  // Planen står i samtalen – ikke på fanen Handlinger – og intet er ændret endnu.
  await expect(tp.page.getByText("Retter ordlyden i første afsnit.")).toBeVisible();
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Gør det med agenten" })).toHaveCount(0);
  // Agent-opgaven blev brugt (JSON-plan med dokumentets afsnit), på dansk.
  const sys = tp.model.lastRequest!.messages[0].content;
  expect(sys).toContain("You are in agent mode");
  expect(sys).toContain("Write summary and reasons in Danish");

  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();

  // Opfølgning: historikken indeholder den første plan og dens udfald.
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Ingen flere ændringer.", actions: [] }) });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Ret det tilbage til det gamle");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Ingen flere ændringer.")).toBeVisible();
  const msgs = tp.model.lastRequest!.messages;
  expect(msgs.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);
  expect(msgs[1].content).toBe("Ændr begyndelsen af første afsnit");
  expect(msgs[2].content).toContain("Retter ordlyden i første afsnit.");
  expect(msgs[2].content).toContain('"action":"replace"');
  expect(msgs[2].content).toMatch(/Outcome:\n- .*: done/);
  expect(msgs[3].content).toContain("Request: Ret det tilbage til det gamle");

  // Et spørgsmål besvares stadig som almindelig tekst.
  tp.model.reply = () => ({ stream: ["Dokumentet handler om levering."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad handler dokumentet om?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Dokumentet handler om levering.")).toBeVisible();
  expect(tp.model.lastRequest!.messages[0].content).not.toContain("You are in agent mode");
});

test("chat: et ønske, ordgenkendelsen ikke fanger, kan stadig udføres i samtalen, når svaret henviser til agenten (#19)", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  await tp.tab("Chat");
  tp.model.reply = () => ({ stream: ["Det kan MIAI's agent gøre direkte i dokumentet: den kan stille priserne op som punkter."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Kan priserne stilles pænere op?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText(/stille priserne op som punkter/)).toBeVisible();

  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Gør prisen til et punkt.", actions: [{ action: "setStyle", paragraph: 2, style: "bullet", reason: "Punkt" }] }),
  });
  await tp.page.getByRole("button", { name: "Gør det med agenten" }).click();
  // Planen kommer i samtalen (fanen skifter ikke), og brugerens besked gentages ikke.
  await expect(tp.page.getByText("Gør prisen til et punkt.")).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Udfør 1 handling" })).toBeVisible();
  await expect(tp.page.getByText("Kan priserne stilles pænere op?")).toHaveCount(1);
  const last = tp.model.lastRequest!.messages.at(-1)!.content;
  expect(last).toContain("Request: Kan priserne stilles pænere op?");
  expect(last).toContain("<material>");
});

test("Excel-chat: en graf bliver et rigtigt Excel-diagram gennem agenten, ikke en tegning i chatten", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Aktier",
      values: [
        ["Aktie", "Kurs"],
        ["NOVO-B", 274.3],
      ],
    },
  });
  await tp.tab("Chat");
  // "Dan en trendgraf" fanges ikke af ordgenkendelsen, så chatten svarer. Vejledningen beder den
  // henvise til agenten og give tallene som en tabel i stedet for at tegne med Mermaid.
  const answer =
    "MIAI's agent kan indsætte grafen som et Excel-diagram i arket.\n\n| Dato | NOVO-B |\n|---|---|\n| 2026-11-23 | 266.35 |\n| 2026-11-24 | 274.65 |\n| 2026-11-25 | 274.30 |";
  tp.model.reply = () => ({ stream: [answer] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Dan en trendgraf for aktien de seneste dage");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText(/indsætte grafen som et Excel-diagram/)).toBeVisible();
  expect(tp.model.lastRequest!.messages[0].content).toContain("must be a real Excel chart, never a drawing");

  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Skriver kurserne i arket og laver et kurvediagram.",
      actions: [
        {
          action: "setValues",
          address: "D1",
          values: [
            ["Dato", "NOVO-B"],
            ["2026-11-23", 266.35],
            ["2026-11-24", 274.65],
            ["2026-11-25", 274.3],
          ],
          reason: "Kurserne fra chatten",
        },
        { action: "addChart", source: "D1:E4", type: "line", title: "NOVO-B", reason: "Udviklingen over tid" },
      ],
    }),
  });
  await tp.page.getByRole("button", { name: "Gør det med agenten" }).click();
  await expect(tp.page.getByText("Skriver kurserne i arket og laver et kurvediagram.")).toBeVisible();
  // Agenten får chatsvaret med tabellen som materiale og ved, at tallene først skal skrives i arket.
  const request = tp.model.lastRequest!;
  expect(request.messages.at(-1)!.content).toContain("| 2026-11-24 | 274.65 |");
  expect(request.messages[0].content).toContain("first write them with setValues");

  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("2 af 2 handlinger udført.")).toBeVisible();
  const calls = await tp.calls();
  expect(calls.find((c) => c.kind === "values")).toMatchObject({ address: "Aktier!D1:E4" });
  expect(calls.find((c) => c.kind === "chartAdd")).toMatchObject({ source: "D1:E4" });
});

test("Excel-chat: tegner modellen alligevel en graf med Mermaid, tilbydes et rigtigt Excel-diagram", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: { sheet: "Aktier", values: [["Aktie"], ["NOVO-B"]] } });
  await tp.tab("Chat");
  const drawing = 'Her er grafen:\n\n```mermaid\nxychart-beta\n    x-axis ["24. nov", "25. nov"]\n    line [274.65, 274.3]\n```';
  tp.model.reply = () => ({ stream: [drawing] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Dan en trendgraf for aktien");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Vil du have grafen som et rigtigt Excel-diagram?")).toBeVisible();

  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Laver et kurvediagram af kurserne.", actions: [] }) });
  await tp.page.getByRole("button", { name: "Gør det med agenten" }).click();
  await expect(tp.page.getByText("Laver et kurvediagram af kurserne.")).toBeVisible();
  // Agenten får tegningens tal som materiale.
  expect(tp.model.lastRequest!.messages.at(-1)!.content).toContain("line [274.65, 274.3]");
});

test("Word-chat: en Mermaid-graf er stadig en figur, ikke et tilbud om et Excel-diagram", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  await tp.tab("Chat");
  tp.model.reply = () => ({ stream: ['```mermaid\nxychart-beta\n    x-axis ["a", "b"]\n    line [1, 2]\n```'] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Vis udviklingen som en graf");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByRole("button", { name: "Kopiér" }).first()).toBeVisible();
  await expect(tp.page.getByText("Vil du have grafen som et rigtigt Excel-diagram?")).toHaveCount(0);
});

test("agentens plan i Word (#22): 'Vis i dokumentet' markerer teksten eller afsnittet, handlingen gælder", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Retter prisen og gør noten til en overskrift.",
      actions: [
        { action: "replace", find: "499 kr", replace: "599 kr", reason: "" },
        { action: "setStyle", paragraph: 5, style: "heading2", reason: "" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  const show = tp.page.getByRole("button", { name: "Vis i dokumentet" });
  await expect(show).toHaveCount(2);
  await show.nth(0).click();
  await show.nth(1).click();
  await expect
    .poll(() =>
      tp.page.evaluate(() =>
        (window as unknown as { __miai: { calls: { kind: string; text?: string; paragraph?: number }[] } }).__miai.calls
          .filter((c) => c.kind === "select")
          .map((c) => ({ text: c.text, paragraph: c.paragraph })),
      ),
    )
    .toEqual([{ text: "499 kr" }, { text: "Gammel note, der skal væk.", paragraph: 5 }]);
});

const TABEL_DOK = [
  { text: "Priser", style: "Heading1" },
  { text: "Vare", cell: [1, 1] as [number, number] },
  { text: "Pris", cell: [1, 2] as [number, number] },
  { text: "Æble", cell: [2, 1] as [number, number] },
  { text: "4 kr", cell: [2, 2] as [number, number] },
  { text: "Pære", cell: [3, 1] as [number, number] },
  { text: "5 kr", cell: [3, 2] as [number, number] },
  { text: "Slut", style: "Normal" },
];
type Tabel = { rows: string[][]; style?: string; styleBuiltIn?: string; headerRowCount?: number };
const tables = (tp: { page: import("@playwright/test").Page }) =>
  tp.page.evaluate(() => (window as unknown as { __miai: { tables: () => Tabel[] } }).__miai.tables());

test("Word-tabeller (#24): ny tabel, ny række, slettet række og rettet celle; fortryd gendanner dokumentet", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: TABEL_DOK });
  const before = await doc(tp);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Opdaterer pristabellen og tilføjer en tabel over butikker.",
      actions: [
        { action: "addTableRow", paragraph: 4, values: ["Banan", "3 kr"], reason: "Ny vare" },
        { action: "deleteTableRow", paragraph: 6, reason: "Udgået" },
        { action: "replace", find: "4 kr", replace: "6 kr", reason: "Ny pris" },
        {
          action: "insertTable",
          position: "after",
          paragraph: 8,
          rows: [
            ["Butik", "By"],
            ["Nord", "Aalborg"],
          ],
          reason: "Butikker",
        },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  // Modellen ser cellerne med tabel, række og kolonne.
  const prompt = (await tp.model.waitForRequest()).messages.map((m) => m.content).join("\n");
  expect(prompt).toContain("¶4 [Normal] (table 1, row 2, col 1) Æble\n¶5 [Normal] (table 1, row 2, col 2) 4 kr");
  expect(prompt).toContain("insertTable {position, paragraph?, rows}");
  await expect(tp.page.getByRole("checkbox", { name: "Ny tabelrække ved ¶4 “Æble”: Banan | 3 kr" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Slet tabelrækken med ¶6 “Pære”" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Indsæt tabel (2 rækker × 2 kolonner): Butik | By" })).toBeChecked();

  await tp.page.getByRole("button", { name: "Udfør 4 handlinger" }).click();
  await expect(tp.page.getByText("4 af 4 handlinger udført.")).toBeVisible();
  expect(await tables(tp)).toEqual([
    {
      style: "Tabelgitter",
      rows: [
        ["Vare", "Pris"],
        ["Æble", "6 kr"],
        ["Banan", "3 kr"],
      ],
    },
    // Den nye tabel får samme typografi som dokumentets tabel og en overskriftsrække.
    {
      style: "Tabelgitter",
      headerRowCount: 1,
      rows: [
        ["Butik", "By"],
        ["Nord", "Aalborg"],
      ],
    },
  ]);
  expect((await doc(tp)).at(-1)!.text).toBe("Aalborg");

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await doc(tp)).toEqual(before);
});

test("Word-tabeller (#24): uden en tabel i dokumentet bruges en indbygget tabeltypografi, og rækkehandlinger kræver en tabel", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Tilføjer en tabel.",
      actions: [
        {
          action: "insertTable",
          position: "end",
          rows: [
            ["Plan", "Pris"],
            ["Basis", "499 kr"],
            ["Plus", "799 kr"],
          ],
          reason: "Prisoversigt",
        },
        { action: "addTableRow", paragraph: 2, values: ["x"], reason: "" },
        { action: "insertTable", position: "end", rows: [["A", "B"], ["kun én"]], reason: "" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Afsnittet står ikke i en tabel")).toBeVisible();
  await expect(tp.page.getByText("Ufuldstændig handling")).toBeVisible();
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
  expect(await tables(tp)).toEqual([
    {
      styleBuiltIn: "GridTable4_Accent1",
      headerRowCount: 1,
      rows: [
        ["Plan", "Pris"],
        ["Basis", "499 kr"],
        ["Plus", "799 kr"],
      ],
    },
  ]);
});

test("kodetilstand i Word (#23): scriptet kører i sandkassen, kan fortrydes, og et fejlende script rulles tilbage", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS, settings: { agentScripts: true } });
  const script = (code: string, description = "Tilføj en linje") => ({
    json: JSON.stringify({ summary: "Script.", actions: [{ action: "runScript", description, code, reason: "" }] }),
  });

  // Excel-objektet findes ikke i Word, og netværk er spærret som i Excel.
  tp.model.reply = () => script("const s = Excel.run; await fetch('https://evil.example');", "Forkert vært");
  await tp.run("Agent: redigér dokumentet");
  expect((await tp.model.waitForRequest()).messages[0].content).toContain("`context` (Word.RequestContext) and `Word` in scope");
  await expect(tp.page.getByText("Bruger noget, sandkassen ikke tillader")).toBeVisible();

  tp.model.reply = () =>
    script('context.document.body.insertParagraph("Tilføjet af script", "End");\nawait context.sync();\nreturn "Tilføjede en linje sidst i dokumentet";');
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByText("Vis koden").click();
  await expect(tp.page.getByText('insertParagraph("Tilføjet af script"')).toBeVisible();
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("Tilføjede en linje sidst i dokumentet", { exact: true })).toBeVisible();
  expect((await doc(tp)).at(-1)!.text).toBe("Tilføjet af script");
  await expect(tp.page.getByText(/Fortryd sætter dokumentets brødtekst tilbage/)).toBeVisible();
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await doc(tp)).toEqual(ORIGINAL);

  // Et script, der ændrer og så fejler: dokumentet er som før, og modellen kan rette det.
  tp.model.reply = () => script('context.document.body.insertParagraph("Halvt færdig", "End");\nawait context.sync();\nthrow new Error("Bum i Word");');
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText(/Bum i Word – ændringerne er rullet tilbage/)).toBeVisible();
  expect(await doc(tp)).toEqual(ORIGINAL);
  await expect(tp.page.getByRole("button", { name: "Bed modellen rette scriptet" })).toBeVisible();
});

test("chatten bevares ved faneskift: svaret skrives færdigt, planen kan stadig udføres, og kladden er der (#26)", async ({ open }) => {
  // Uden dokumentadresse huskes intet i lageret: alt, der er der bagefter, kommer fra den kørende chat.
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.6" }, paragraphs: PARAS });
  const field = tp.page.getByPlaceholder("Spørg om dit dokument…");
  await tp.tab("Chat");

  // Et langt svar, der skrives, mens brugeren kigger i Indstillinger.
  // Teksten må ikke gentage sig selv: så afbryder vagten mod gentagelsessløjfer svaret.
  const long = `Første del af svaret. ${Array.from({ length: 40 }, (_, i) => `Punkt ${i + 1} gælder ${i * 7 + 3} sider.`).join(" ")} Sidste sætning i svaret.`;
  tp.model.reply = () => ({ stream: chunks(long, 3) });
  await field.fill("Hvad handler dokumentet om?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText(/Første del af svaret/)).toBeVisible();
  await tp.tab("Indstillinger");
  await expect(tp.page.getByRole("button", { name: "Gem" })).toBeVisible();
  await tp.tab("Chat");
  await expect(tp.page.getByText(/Sidste sætning i svaret\./)).toBeVisible();
  await expect(tp.page.getByText("Hvad handler dokumentet om?")).toBeVisible();

  // En plan i samtalen kan stadig udføres efter et skift til Handlinger og tilbage.
  const first = PARAS[1].text.split(" ").slice(0, 2).join(" ");
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Retter ordlyden i første afsnit.",
      actions: [{ action: "replace", find: first, replace: "Ny begyndelse", reason: "Ønsket" }],
    }),
  });
  await field.fill("Ændr begyndelsen af første afsnit");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByRole("button", { name: "Udfør 1 handling" })).toBeVisible();
  await field.fill("Kladde til næste spørgsmål");
  await tp.tab("Handlinger");
  await tp.tab("Chat");
  await expect(field).toHaveValue("Kladde til næste spørgsmål");
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
});
