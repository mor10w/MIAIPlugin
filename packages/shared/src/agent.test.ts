import { describe, expect, it } from "vitest";
import { z } from "zod";
import {
  AgentPlan,
  agentToolGuide,
  cellCount,
  checkPlan,
  readRequests,
  isDestructive,
  MAX_AGENT_ACTIONS,
  numberedRows,
  rangeBox,
  renderExcelStructure,
  renderParagraphs,
  sheetOfAddress,
} from "./agent.js";
import { minimalEdit, occurrences, parseBlocks, parseInline } from "./agentText.js";
import { TaskRequest } from "./types.js";
import { TASKS } from "./tasks.js";

const agent = TASKS.find((t) => t.id === "agent")!;
const doc = "Vi har mange kunder i Jylland. Kunden betaler ved levering.\nKontakt os på info@firma.dk.";

describe("agent-plan (#6)", () => {
  it("skemaet kan omsættes til JSON Schema (json_schema-tilstand) og har et loft over antal handlinger", () => {
    const js = JSON.stringify(z.toJSONSchema(AgentPlan));
    expect(js).toContain('"replace"');
    expect(js).toContain('"setFormula"');
    expect(AgentPlan.safeParse({ summary: "", actions: Array(MAX_AGENT_ACTIONS + 1).fill({ action: "delete", find: "x", reason: "" }) }).success).toBe(false);
  });

  it("ukendte handlinger (fx netværk eller filer) afvises af skemaet", () => {
    expect(AgentPlan.safeParse({ summary: "", actions: [{ action: "fetchUrl", url: "https://evil.example", reason: "" }] }).success).toBe(false);
    expect(AgentPlan.safeParse({ summary: "", actions: [{ action: "setFont", find: "x", font: "Comic Sans", reason: "" }] }).success).toBe(false);
  });

  it("checkPlan (Word): forkert vært, tekst der ikke findes, afsnitsnumre og tabeller", () => {
    const paragraphs = [
      { text: "Kundeinformation", style: "Heading1" },
      { text: "Vi har mange kunder i Jylland. Kunden betaler ved levering.", style: "Normal" },
      { text: "Pris", style: "Normal", table: true },
      { text: "Kontakt os på info@firma.dk.", style: "Normal" },
    ];
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "replace", find: "kunder", replace: "klienter", all: true, reason: "" },
        { action: "replace", find: "findes ikke", replace: "x", reason: "" },
        { action: "insert", text: "## Levering", position: "after", paragraph: 2, reason: "" },
        { action: "insert", text: "Ny", position: "after", paragraph: 99, reason: "" },
        { action: "delete", paragraph: 3, reason: "" },
        { action: "delete", reason: "" },
        { action: "setStyle", paragraph: 4, style: "quote", reason: "" },
        { action: "move", paragraph: 4, to: 2, position: "before", reason: "" },
        { action: "replace", find: "Pris", replace: "Pris (kr.)", reason: "" },
        { action: "setFormula", address: "F2", formula: "=1", reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "word", text: "", paragraphs }).map((c) => c.problem)).toEqual([
      undefined,
      "not-found",
      undefined,
      "bad-paragraph",
      "in-table",
      "invalid",
      undefined,
      undefined,
      undefined, // tekst i en tabelcelle må gerne rettes med replace
      "wrong-host",
    ]);
  });

  it("checkPlan uden afsnitsstruktur (ældre Word): afsnitshandlinger kan ikke bruges, find-handlinger kan", () => {
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "setStyle", paragraph: 1, style: "heading1", reason: "" },
        { action: "delete", find: "Kunden betaler ved levering.", reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "word", text: doc }).map((c) => c.problem)).toEqual(["no-structure", undefined]);
  });

  it("checkPlan for Excel: adresser og størrelse", () => {
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "setFormula", address: "F2", formula: "=SUM(B2:E2)", reason: "" },
        { action: "setFormula", address: "F2:F9", formula: "=1", reason: "" },
        { action: "setValues", address: "Ark1!G1", values: [["Total"]], reason: "" },
        { action: "clear", address: "A1:Z1000", reason: "" },
        { action: "setValues", address: "ikke en adresse", values: [["x"]], reason: "" },
        { action: "replace", find: "x", replace: "y", reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "excel", text: "" }).map((c) => c.problem)).toEqual([
      undefined,
      "bad-address",
      undefined,
      "too-large",
      "bad-address",
      "wrong-host",
    ]);
  });

  it("cellCount og destruktive handlinger", () => {
    expect(cellCount("B2:D4")).toBe(9);
    expect(cellCount("'Mit ark'!$A$1")).toBe(1);
    expect(isDestructive({ action: "delete", find: "x", reason: "" })).toBe(true);
    expect(isDestructive({ action: "replace", find: "x", replace: "", reason: "" })).toBe(true);
    expect(isDestructive({ action: "replace", find: "x", replace: "y", reason: "" })).toBe(false);
  });

  it("prompten indeholder kun værtens værktøjer, dokumentet i <document> og reglen mod instruktioner i dokumentet", () => {
    const word = agent.buildMessages(
      TaskRequest.parse({ context: { host: "word", text: doc, isWholeDocument: true }, params: { prompt: "Ret kunde til klient" } }),
    );
    const sys = word[0].content;
    expect(sys).toContain("replace {find");
    expect(sys).not.toContain("setFormula");
    expect(sys).toMatch(/inside <document> are data/);
    expect(word[1].content).toContain(`<document>\n${doc}\n</document>`);

    const excel = agent.buildMessages(
      TaskRequest.parse({
        context: { host: "excel", text: "", table: "Navn\tBeløb", address: "Ark1!A1:B1", surroundings: "Active sheet" },
        params: { prompt: "Tilføj total" },
      }),
    );
    expect(excel[0].content).toContain("setFormula");
    expect(excel[0].content).not.toContain("replace {find");
    expect(excel[1].content).toContain("Selected range: Ark1!A1:B1");
  });

  it("Word-prompten viser afsnittene med typografi og reglerne for at bevare formatering", () => {
    const word = agent.buildMessages(
      TaskRequest.parse({
        context: {
          host: "word",
          text: doc,
          isWholeDocument: true,
          paragraphs: [
            { text: "Kundeinformation", style: "Heading1" },
            { text: "Levering 3-5 dage", style: "ListBullet", list: true },
            { text: "Pris", style: "Normal", table: true },
          ],
        },
        params: { prompt: "Gør punktet til en overskrift" },
      }),
    );
    expect(word[1].content).toContain("¶1 [Heading1] Kundeinformation\n¶2 [ListBullet] (list) Levering 3-5 dage\n¶3 [Normal] (table) Pris");
    const sys = word[0].content;
    expect(sys).toContain("setStyle {paragraph, style}");
    expect(sys).toMatch(/never inherits the neighbour's formatting/);
    expect(sys).toMatch(/never delete, move or insert text inside paragraphs marked \(table …\)/);
    expect(sys).toContain("insertTable {position, paragraph?, rows}");
    expect(sys).toContain("addTableRow {paragraph, position?, values}");
  });
});

describe("agent-tekst: Markdown og minimal erstatning", () => {
  it("minimalEdit ændrer kun de ord, der er forskellige", () => {
    const cases: [string, string, string][] = [
      ["Pris: 499 kr pr. md.", "Pris: 599 kr pr. md.", "499"],
      ["Vi har mange kunder.", "Vi har mange klienter.", "kunder."],
      ["Ring til Bo.", "Ring til Bo eller Anna.", "Bo."],
      ["a b c", "a c", "b "],
    ];
    for (const [find, replace, oldCore] of cases) {
      const r = minimalEdit(find, replace);
      expect(r.oldCore, find).toBe(oldCore);
      // Sat sammen giver det præcis den ønskede tekst.
      expect(r.before + r.newCore + find.slice(r.before.length + r.oldCore.length)).toBe(replace);
    }
  });
  it("inline-Markdown: fed, kursiv og kun sikre links", () => {
    const r = parseInline("Se **vigtigt**, *kursiv* og [MIAI](https://miai.dk) men ikke [ond](javascript:x)");
    expect(r.text).toBe("Se vigtigt, kursiv og MIAI men ikke ond");
    expect(r.spans).toEqual([
      { start: 3, end: 10, bold: true },
      { start: 12, end: 18, italic: true },
      { start: 22, end: 26, link: "https://miai.dk" },
    ]);
    expect(parseInline("2*3*4 og snake_case_navn").spans).toEqual([]);
  });
  it("blok-Markdown: én linje pr. afsnit med typografi", () => {
    expect(parseBlocks("## Levering\n- 3-5 dage\n1. Bestil\nAlm. tekst\n\n> Citat", "normal").map((b) => [b.style, b.inline.text])).toEqual([
      ["heading2", "Levering"],
      ["bullet", "3-5 dage"],
      ["numbered", "Bestil"],
      ["normal", "Alm. tekst"],
      ["quote", "Citat"],
    ]);
  });
  it("occurrences tæller uden overlap", () => {
    expect(occurrences("aaa", "aa")).toBe(1);
    expect(occurrences("kunder og kunder", "kunder")).toBe(2);
  });
});

describe("chat → agent", () => {
  it("chatsvaret gives til agenten som indhold i <material>", () => {
    const agentTask = TASKS.find((t) => t.id === "agent")!;
    const msgs = agentTask.buildMessages(
      TaskRequest.parse({ context: { host: "word", text: "x" }, params: { prompt: "Tilføj afsnittet", material: "## Levering\nVi leverer på 3 dage." } }),
    );
    expect(msgs[1].content).toContain("<material>\n## Levering\nVi leverer på 3 dage.\n</material>");
  });
});

describe("agent i PowerPoint og Excel", () => {
  const slides = [
    {
      layout: "Title Slide",
      shapes: [
        { name: "Title 1", kind: "title" as const, text: "Kvartalsrapport" },
        { name: "Subtitle 2", kind: "body" as const, text: "Q3 2026" },
      ],
    },
    {
      layout: "Title and Content",
      shapes: [
        { name: "Title 1", kind: "title" as const, text: "Resultater" },
        { name: "Content 2", kind: "body" as const, text: "Omsætning +12 %\nMargin 38 %" },
        { name: "Picture 3", kind: "other" as const, text: "" },
      ],
    },
  ];
  it("checkPlan (PowerPoint): figurer, slides og layouts", () => {
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "replace", find: "+12 %", replace: "+14 %", reason: "" },
        { action: "setText", shape: "2.2", text: "- Omsætning +14 %\n- Margin 39 %", reason: "" },
        { action: "setText", shape: "2.3", text: "x", reason: "" },
        { action: "setText", shape: "3.1", text: "x", reason: "" },
        { action: "addSlide", after: 2, layout: "Title and Content", title: "Næste skridt", reason: "" },
        { action: "addSlide", after: 2, layout: "Mit eget layout", title: "x", reason: "" },
        { action: "deleteSlide", slide: 5, reason: "" },
        { action: "moveSlide", slide: 2, to: 1, reason: "" },
        { action: "delete", paragraph: 1, reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "powerpoint", text: "", slides, layouts: ["Title Slide", "Title and Content"] }).map((c) => c.problem)).toEqual([
      undefined,
      undefined,
      "invalid",
      "bad-slide",
      undefined,
      "bad-layout",
      "bad-slide",
      undefined,
      "wrong-host",
    ]);
  });

  it("checkPlan (Excel): appendRows kræver en kendt tabel og højst én værdi pr. kolonne", () => {
    const excel = { tables: [{ name: "Salg", address: "Ark1!A1:C4", headers: ["Region", "Beløb", "Andel"] }], columnFormats: [] };
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "appendRows", table: "Salg", values: [["Fyn", 70, null]], reason: "" },
        { action: "appendRows", table: "Findes", values: [["x"]], reason: "" },
        { action: "appendRows", table: "Salg", values: [["a", 1, 2, 3]], reason: "" },
        { action: "setValues", address: "E2", values: [[1200.5, true, null, "tekst"]], reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "excel", text: "", excel }).map((c) => c.problem)).toEqual([undefined, "bad-table", "invalid", undefined]);
  });

  it("checkPlan (Excel): sletning af rækker og tabelformatering må ikke ødelægge en tabel", () => {
    const excel = { tables: [{ name: "Salg", address: "Ark1!A1:C6", headers: ["Region", "Beløb", "Andel"] }], columnFormats: [] };
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "deleteRows", address: "A3:C4", reason: "" },
        { action: "deleteRows", address: "A1:C2", reason: "" },
        { action: "deleteRows", address: "A6:C8", reason: "" },
        { action: "deleteRows", address: "A10:C12", reason: "" },
        { action: "deleteEmptyRows", target: "Salg", reason: "" },
        { action: "deleteEmptyRows", target: "Findes", reason: "" },
        { action: "deleteEmptyRows", target: "A1:F200", reason: "" },
        { action: "formatTable", target: "Salg", style: "TableStyleMedium9", reason: "" },
        { action: "formatTable", target: "B4:E9", style: "TableStyleLight1", reason: "" },
        { action: "formatTable", target: "E1:G9", style: "TableStyleDark11", reason: "" },
        { action: "setNumberFormat", address: "B2:B6", format: "#,##0", reason: "" },
        { action: "autofitColumns", address: "A1:C6", reason: "" },
        { action: "autofitColumns", address: "Salg", reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "excel", text: "", excel }).map((c) => c.problem)).toEqual([
      undefined,
      "in-table", // overskriften
      "in-table", // stikker ud under tabellen
      undefined,
      undefined,
      "bad-table",
      undefined,
      undefined,
      "in-table", // ny tabel oven i en eksisterende
      undefined,
      undefined,
      undefined,
      "bad-address",
    ]);
    // Kun Excels indbyggede tabeltypografier.
    expect(() => AgentPlan.parse({ summary: "s", actions: [{ action: "formatTable", target: "Salg", style: "Lilla og guld", reason: "" }] })).toThrow();
    expect(isDestructive(plan.actions[0])).toBe(true);
    expect(isDestructive(plan.actions[4])).toBe(true);
  });

  it("checkPlan (Excel): sortering, betinget formatering, diagram og frysning", () => {
    const excel = { tables: [{ name: "Salg", address: "Ark1!A1:C5", headers: ["Region", "Måned", "Omsætning"] }], columnFormats: [] };
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "sortRange", target: "Salg", by: "Omsætning", order: "desc", reason: "" },
        { action: "sortRange", target: "Findes", by: "A", reason: "" },
        { action: "sortRange", target: "A1:C1", by: "A", reason: "" },
        { action: "conditionalFormat", address: "C2:C5", rule: "greaterThan", reason: "" },
        { action: "conditionalFormat", address: "C2:C5", rule: "greaterThan", value: 100, reason: "" },
        { action: "addChart", source: "Salg", type: "pie", reason: "" },
        { action: "freezeRows", rows: 1, reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "excel", text: "", excel }).map((c) => c.problem)).toEqual([
      undefined,
      "bad-table",
      "invalid", // kun en overskrift
      "invalid", // mangler værdi
      undefined,
      undefined,
      undefined,
    ]);
    expect(plan.actions[1]).toMatchObject({ order: "asc" });
  });

  it("checkPlan (Excel): runScript kræver kodetilstand, et sikkert script og at stå alene", () => {
    const script = {
      action: "runScript",
      description: "Pivot",
      code: "const s = context.workbook.worksheets.getActiveWorksheet(); await context.sync();",
      reason: "",
    };
    const alone = AgentPlan.parse({ summary: "s", actions: [script] });
    expect(checkPlan(alone, { host: "excel", text: "" })[0].problem).toBe("scripts-off");
    expect(checkPlan(alone, { host: "excel", text: "" }, { scripts: true })[0].problem).toBeUndefined();
    const unsafe = AgentPlan.parse({ summary: "s", actions: [{ ...script, code: "await fetch('https://x')" }] });
    expect(checkPlan(unsafe, { host: "excel", text: "" }, { scripts: true })[0].problem).toBe("unsafe-script");
    const mixed = AgentPlan.parse({ summary: "s", actions: [script, { action: "freezeRows", rows: 1, reason: "" }] });
    expect(checkPlan(mixed, { host: "excel", text: "" }, { scripts: true }).map((c) => c.problem)).toEqual([undefined, "with-script"]);
    // Kun i Excel.
    expect(checkPlan(alone, { host: "word", text: "" }, { scripts: true })[0].problem).toBe("wrong-host");
    expect(isDestructive(alone.actions[0])).toBe(true);
  });

  it("checkPlan (Excel): saveSheet kræver eksport og validerer filnavnet", () => {
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "saveSheet", fileName: "Svar 2026.xlsx", reason: "" },
        { action: "saveSheet", fileName: "../evil.xlsx", sheet: "Data", reason: "" },
        { action: "saveSheet", fileName: "fil.txt", reason: "" },
      ],
    });
    // Eksport er fra (eller der er ingen bridge): handlingen findes, men kan ikke bruges.
    expect(checkPlan(plan, { host: "excel", text: "" }).map((c) => c.problem)).toEqual(["no-bridge", "no-bridge", "no-bridge"]);
    // Med eksport valideres navnet (ingen stipelementer, .xlsx-påliggende).
    expect(checkPlan(plan, { host: "excel", text: "" }, { export: true }).map((c) => c.problem)).toEqual([undefined, "invalid", "invalid"]);
    // Kun i Excel.
    expect(checkPlan(plan, { host: "word", text: "" }, { export: true })[0].problem).toBe("wrong-host");
    // Ikke destruktiv – den ændrer ikke arbejdsbogs, men skriver en ny fil.
    expect(isDestructive(plan.actions[0])).toBe(false);
  });

  it("prompten: saveSheet vises kun, når agentExport=on", () => {
    const req = (extra: Record<string, string> = {}) =>
      TaskRequest.parse({ context: { host: "excel", text: "", address: "Ark1!A1" }, params: { prompt: "Gem arket som fil", ...extra } });
    expect(agent.buildMessages(req())[0].content).not.toContain("saveSheet {fileName");
    expect(agent.buildMessages(req({ agentExport: "on" }))[0].content).toContain("saveSheet {fileName");
    // Word-prompten får aldrig eksporten.
    const word = agent.buildMessages(
      TaskRequest.parse({ context: { host: "word", text: doc, isWholeDocument: true }, params: { prompt: "s", agentExport: "on" } }),
    );
    expect(word[0].content).not.toContain("saveSheet");
  });

  it("checkPlan (Excel, flere ark #21): arknavne skal findes eller oprettes af planen, og tabeller tjekkes pr. ark", () => {
    const excel = {
      tables: [{ name: "Salg", address: "Ark1!A1:C5", headers: ["Region", "Måned", "Omsætning"] }],
      columnFormats: [],
      activeSheet: "Ark1",
      sheets: [
        { name: "Ark1", used: "A1:C5", headers: ["Region", "Måned", "Omsætning"] },
        { name: "Mit ark", headers: [] },
      ],
    };
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "addSheet", name: "Oversigt", reason: "" },
        { action: "setValues", address: "Oversigt!A1", values: [["x"]], reason: "" },
        { action: "setFormula", address: "'Mit ark'!B2", formula: "=SUM(Ark1!C2:C5)", reason: "" },
        { action: "setValues", address: "Findes!A1", values: [["x"]], reason: "" },
        { action: "addSheet", name: "ark1", reason: "" },
        { action: "renameSheet", sheet: "Mit ark", name: "Noter", reason: "" },
        { action: "renameSheet", sheet: "Væk", name: "Nyt", reason: "" },
        // Samme celler som tabellens overskrift – men på et andet ark, så det er i orden.
        { action: "deleteRows", address: "'Mit ark'!A1:C2", reason: "" },
        { action: "deleteRows", address: "A1:C2", reason: "" },
        { action: "formatTable", target: "Oversigt!A1:B4", style: "TableStyleLight9", reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "excel", text: "", excel }).map((c) => c.problem)).toEqual([
      undefined,
      undefined,
      undefined,
      "bad-sheet",
      "bad-sheet",
      undefined,
      "bad-sheet",
      undefined,
      "in-table",
      undefined,
    ]);
    expect(() => AgentPlan.parse({ summary: "s", actions: [{ action: "addSheet", name: "A/B", reason: "" }] })).toThrow();
    expect(sheetOfAddress("'Mit ark'!B2")).toBe("Mit ark");
    expect(sheetOfAddress("B2")).toBeUndefined();
    expect(renderExcelStructure(excel)).toContain("Sheets:\n- Ark1 (active): A1:C5, headers: Region | Måned | Omsætning\n- Mit ark: empty");
  });

  it("readRange (#20): læsninger udføres af add-in'et før planen, aldrig som handling, og kun gyldige udføres", () => {
    const excel = {
      tables: [],
      columnFormats: [],
      activeSheet: "Ark1",
      sheets: [
        { name: "Ark1", headers: [] },
        { name: "Budget", headers: [] },
      ],
    };
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        { action: "readRange", address: "Budget!A1:F40", reason: "" },
        { action: "readRange", address: "Findes!A1:B2", reason: "" },
        { action: "readRange", address: "A1:Z1000", reason: "" },
        { action: "readRange", address: "ikke en adresse", reason: "" },
      ],
    });
    const ctx = { host: "excel" as const, text: "", excel };
    expect(readRequests(plan, ctx).map((r) => r.address)).toEqual(["Budget!A1:F40"]);
    expect(readRequests(plan, { host: "word", text: "" })).toEqual([]);
    // Står en læsning tilbage i en plan, kan den ikke vælges.
    expect(checkPlan(plan, ctx).map((c) => c.problem)).toEqual(["invalid", "bad-sheet", "too-large", "bad-address"]);
    expect(agentToolGuide("excel", { reads: true })).toContain("readRange {address}");
    expect(agentToolGuide("excel")).not.toContain("readRange {address}");
    expect(agentToolGuide("word", { reads: true })).not.toContain("readRange");
  });

  it("agentens prompt (#20): det læste står som data, og en delvis fejlet plan rettes uden at gentage det udførte", () => {
    const msgs = agent.buildMessages(
      TaskRequest.parse({
        context: { host: "excel", text: "", address: "Ark1!A1:B2", observations: "Budget!A1:B2:\n1\tRegion\tBudget" },
        params: { prompt: "Sammenlign med budgettet", agentReads: "on", previousError: "- Sortér: failed", previousCode: "[]", previousPartial: "true" },
      }),
    );
    expect(msgs[0].content).toContain("readRange {address}");
    expect(msgs[1].content).toContain("Cells you asked to read (you have these now – do not ask for them again):\n<document>\nBudget!A1:B2:");
    expect(msgs[1].content).toContain("return ONLY actions that fix what failed");
  });

  it("checkPlan (Excel #25): pivottabellens felter og datavalideringens regel", () => {
    const excel = {
      tables: [{ name: "Salg", address: "Ark1!A1:C5", headers: ["Region", "Måned", "Omsætning"] }],
      columnFormats: [],
      activeSheet: "Ark1",
      sheets: [
        { name: "Ark1", used: "A1:C5", headers: ["Region", "Måned", "Omsætning"] },
        { name: "Data", used: "B2:D9", headers: ["Kunde", "Antal", "Pris"] },
      ],
    };
    const pivot = { action: "addPivotTable", source: "Salg", rows: ["Region"], values: [{ field: "omsætning" }], destination: "F1", reason: "" };
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        pivot,
        { ...pivot, rows: ["Land"] },
        { ...pivot, source: "Data!B2:D9", rows: ["Kunde"], values: [{ field: "Pris", summarize: "max" }] },
        { ...pivot, source: "Data!B2:D9", rows: ["Region"] },
        // Midt i arket kendes overskrifterne ikke – Excel afgør det.
        { ...pivot, source: "Data!B4:D9", rows: ["Hvad som helst"], values: [{ field: "x" }] },
        { ...pivot, destination: "F1:G2" },
        { ...pivot, columns: ["Region"] },
        { ...pivot, destination: "Findes!A1" },
        { action: "setValidation", address: "B2:B5", list: ["Ja", "Nej"], reason: "" },
        { action: "setValidation", address: "B2:B5", list: ["Ja"], min: 1, reason: "" },
        { action: "setValidation", address: "B2:B5", reason: "" },
        { action: "setValidation", address: "B2:B5", min: 5, max: 1, reason: "" },
        { action: "setValidation", address: "B2:B5", max: 100, whole: true, reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "excel", text: "", excel }).map((c) => c.problem)).toEqual([
      undefined,
      "bad-field",
      undefined,
      "bad-field",
      undefined,
      "bad-address",
      "invalid",
      "bad-sheet",
      undefined,
      "invalid",
      "invalid",
      "invalid",
      undefined,
    ]);
    expect(plan.actions[0]).toMatchObject({ columns: [], values: [{ field: "omsætning", summarize: "sum" }] });
  });

  it("Word-tabeller (#24): celler vises med række og kolonne, og rækkehandlinger kræver en tabel", () => {
    const paragraphs = [
      { text: "Priser", style: "Heading1" },
      { text: "Vare", style: "Normal", table: true, cell: { table: 1, row: 1, col: 1 } },
      { text: "Pris", style: "Normal", table: true, cell: { table: 1, row: 1, col: 2 } },
      { text: "Æble", style: "Normal", table: true, cell: { table: 1, row: 2, col: 1 } },
      { text: "4 kr", style: "Normal", table: true, cell: { table: 1, row: 2, col: 2 } },
      { text: "Én række", style: "Normal", table: true, cell: { table: 2, row: 1, col: 1 } },
    ];
    expect(renderParagraphs(paragraphs)).toContain("¶2 [Normal] (table 1, row 1, col 1) Vare\n¶3 [Normal] (table 1, row 1, col 2) Pris");
    const plan = AgentPlan.parse({
      summary: "s",
      actions: [
        {
          action: "insertTable",
          position: "after",
          paragraph: 1,
          rows: [
            ["A", "B"],
            ["1", "2"],
          ],
          reason: "",
        },
        { action: "insertTable", position: "after", paragraph: 2, rows: [["A"], ["1"]], reason: "" },
        { action: "insertTable", position: "end", rows: [["A", "B"], ["1"]], reason: "" },
        { action: "addTableRow", paragraph: 4, values: ["Pære", "5 kr"], reason: "" },
        { action: "addTableRow", paragraph: 4, values: ["Pære", "5 kr", "for mange"], reason: "" },
        { action: "addTableRow", paragraph: 1, values: ["x"], reason: "" },
        { action: "deleteTableRow", paragraph: 5, reason: "" },
        { action: "deleteTableRow", paragraph: 6, reason: "" },
        { action: "deleteTableRow", paragraph: 99, reason: "" },
      ],
    });
    expect(checkPlan(plan, { host: "word", text: "", paragraphs }).map((c) => c.problem)).toEqual([
      undefined,
      "in-table",
      "invalid",
      undefined,
      "invalid",
      "not-in-table",
      undefined,
      "invalid", // tabellens eneste række
      "bad-paragraph",
    ]);
    expect(plan.actions[3]).toMatchObject({ position: "after" });
    expect(isDestructive(plan.actions[6])).toBe(true);
    expect(() => AgentPlan.parse({ summary: "s", actions: [{ action: "insertTable", position: "end", rows: [["kun overskrift"]], reason: "" }] })).toThrow();
  });

  it("rangeBox og numberedRows: arkets rækkenumre foran de markerede værdier", () => {
    expect(rangeBox("Ark1!$B$2:D4")).toEqual({ r0: 1, c0: 1, r1: 3, c1: 3 });
    expect(rangeBox("AA10")).toEqual({ r0: 9, c0: 26, r1: 9, c1: 26 });
    expect(numberedRows("Navn\tBeløb\nAnna\t1", "Ark1!A4:B5")).toBe("4\tNavn\tBeløb\n5\tAnna\t1");
    expect(numberedRows("x", undefined)).toBe("x");
  });

  it("samtalens historik går med til agenten, mellem systemprompten og det nye ønske (#19)", () => {
    const msgs = agent.buildMessages(
      TaskRequest.parse({
        context: { host: "excel", text: "", table: "Navn\tBeløb", address: "Ark1!A1:B1" },
        params: { prompt: "Nej, faldende i stedet" },
        history: [
          { role: "user", content: "Sortér efter beløb" },
          { role: "assistant", content: "Sorterer stigende.\nOutcome:\n- Sortér Ark1 efter Beløb (stigende): done" },
        ],
      }),
    );
    expect(msgs.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);
    expect(msgs[2].content).toContain("Outcome:");
    expect(msgs[3].content).toContain("Request: Nej, faldende i stedet");
  });

  it("prompten viser slides med layout og figurer, og Excels tabeller og talformater", () => {
    const ppt = agent.buildMessages(
      TaskRequest.parse({
        context: { host: "powerpoint", text: "", slides, layouts: ["Title Slide", "Title and Content"] },
        params: { prompt: "Opdater tallene" },
      }),
    );
    expect(ppt[1].content).toContain(
      "Slide 2 (layout: Title and Content)\n  2.1 [title] Title 1: Resultater\n  2.2 [body] Content 2: Omsætning +12 % ⏎ Margin 38 %",
    );
    expect(ppt[1].content).toContain("Layouts: Title Slide | Title and Content");
    expect(ppt[0].content).toContain("addSlide {after, layout?, title, body?}");
    expect(ppt[0].content).not.toContain("setFormula");

    const xl = agent.buildMessages(
      TaskRequest.parse({
        context: {
          host: "excel",
          text: "",
          address: "Ark1!A1",
          excel: { tables: [{ name: "Salg", address: "Ark1!A1:C4", headers: ["Region", "Beløb", "Andel"] }], columnFormats: ["B: #,##0 kr", "C: 0%"] },
        },
        params: { prompt: "Tilføj Fyn" },
      }),
    );
    expect(xl[1].content).toContain("- Salg (Ark1!A1:C4): Region | Beløb | Andel\nNumber formats: B: #,##0 kr; C: 0%");
    expect(xl[0].content).toMatch(/appendRows \{table, values\}/);
  });
});
