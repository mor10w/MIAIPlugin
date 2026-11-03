import { expect, test } from "./fixtures/test";

/** Agent-tilstand (#6): plan → bekræft → udfør, med fortryd af hele sessionen. */

const DOC = "Vi har mange kunder i Jylland. Alle kunder får besked.\nKontakt os på info@firma.dk.\nGammel note, der skal væk.";
const EDITS = (kinds: string[]) => (c: { kind: string }) => kinds.includes(c.kind);
const DOC_EDITS = EDITS(["replace", "insertText", "insertParagraph", "delete", "comment", "style", "insertOoxml"]);

test("Word: planen vises først, intet ændres før Udfør; valgte handlinger udføres som sporede ændringer og kan fortrydes", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, document: DOC });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Retter kunder til klienter, tilføjer en overskrift og fjerner den gamle note.",
      actions: [
        { action: "replace", find: "kunder", replace: "klienter", all: true, reason: "Ønsket ordvalg" },
        { action: "insert", text: "Kundeinformation", position: "start", style: "heading1", reason: "Overskrift" },
        { action: "delete", find: "Gammel note, der skal væk.", reason: "Forældet" },
        { action: "comment", find: "info@firma.dk", text: "Er adressen stadig gyldig?", reason: "Tjek" },
        { action: "replace", find: "findes ikke i teksten", replace: "x", reason: "Opdigtet" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();

  // Hele dokumentet sendes (ingen markering), indpakket som data.
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain(`<document>\n${DOC}\n</document>`);

  // Planen: læsbare handlinger, destruktive fremhævet, tekst der ikke findes, er fravalgt.
  await expect(tp.page.getByRole("checkbox", { name: "Erstat alle “kunder” med “klienter”" })).toBeChecked();
  await expect(tp.page.getByText("Fjerner indhold")).toBeVisible();
  const missing = tp.page.getByRole("checkbox", { name: "Erstat “findes ikke i teksten” med “x”" });
  await expect(missing).toBeDisabled();
  await expect(tp.page.getByText("Teksten findes ikke i dokumentet")).toBeVisible();
  expect((await tp.calls()).filter(DOC_EDITS)).toEqual([]);

  // Fravælg kommentaren og udfør resten.
  await tp.page.getByRole("checkbox", { name: /Kommentar på “info@firma.dk”/ }).uncheck();
  await expect(tp.page.getByRole("switch", { name: "Som sporede ændringer" })).toBeChecked();
  await tp.page.getByRole("button", { name: "Udfør 3 handlinger" }).click();
  await expect(tp.page.getByText("3 af 3 handlinger udført.")).toBeVisible();

  const calls = await tp.calls();
  expect(calls.filter((c) => c.kind === "replace")).toEqual([
    expect.objectContaining({ from: "kunder", to: "klienter" }),
    expect.objectContaining({ from: "kunder", to: "klienter" }),
  ]);
  expect(calls).toContainEqual(expect.objectContaining({ kind: "insertParagraph", text: "Kundeinformation", loc: "Start" }));
  expect(calls).toContainEqual(expect.objectContaining({ kind: "style", style: "Heading1" }));
  expect(calls).toContainEqual(expect.objectContaining({ kind: "delete", text: "Gammel note, der skal væk." }));
  expect(calls.some((c) => c.kind === "comment")).toBe(false);
  await expect(tp.page.getByText("2 steder")).toBeVisible();

  // Fortryd gendanner øjebliksbilledet fra før sessionen.
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await tp.find("insertOoxml")).toMatchObject({ ooxml: "<pkg:package>øjebliksbillede</pkg:package>", loc: "Replace" });
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
