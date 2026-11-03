import { describe, expect, it } from "vitest";
import { z } from "zod";
import { AgentPlan, cellCount, checkPlan, isDestructive, MAX_AGENT_ACTIONS } from "./agent.js";
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
    expect(sys).toMatch(/never delete, move or insert inside paragraphs marked \(table\)/);
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
