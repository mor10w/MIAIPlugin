import type { Page } from "@playwright/test";
import { expect, test } from "./fixtures/test";

/**
 * Formelsporing i Excel (#76): add-in'et gennemløber Excels afhængighedsgraf og giver modellen en
 * kompakt liste. Find fejlkilden, Forklar formel og agentens læseløkke bruger den.
 */

/** En #DIV/0!-kæde over to ark (tre med Data): Resultat!B3 ← B2 ← Beregning!B5 (=B3/B4) ← B4 (= 0). */
const DIV = {
  sheet: "Resultat",
  values: [
    ["Nøgletal", "Værdi"],
    ["Pris pr. stk. x 2", "=Beregning!B5*2"],
    ["Med gebyr", "=B2+1"],
  ] as (string | number)[][],
  results: { B2: "#DIV/0!", B3: "#DIV/0!" },
  sheets: [
    {
      sheet: "Beregning",
      values: [
        ["Post", "Beløb"],
        ["", ""],
        ["Omsætning", 1200],
        ["Antal", "=Data!B2-Data!B3"],
        ["Pris pr. stk.", "=B3/B4"],
      ] as (string | number)[][],
      results: { B4: 0, B5: "#DIV/0!" },
    },
    {
      sheet: "Data",
      values: [
        ["Lager", "Stk."],
        ["Ind", 5],
        ["Ud", 5],
      ] as (string | number)[][],
    },
  ],
};

const select = (page: Page, address: string) =>
  page.evaluate((a) => (window as unknown as { __miai: { selectRange: (x: string) => void } }).__miai.selectRange(a), address);
const prompt = (tp: { model: { lastRequest?: { messages: { content: string }[] } } }) => tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
/** De sporings-API'er, adapteren har kaldt, fx "getDirectPrecedents Resultat!B3". */
const traces = async (tp: { calls: () => Promise<{ kind: string; [k: string]: unknown }[]> }) =>
  (await tp.calls()).filter((c) => c.kind === "trace").map((c) => `${String(c.method)} ${String(c.address)}`);

const ANSWER =
  "**Kilden:** fejlen opstår i Beregning!B5 (`=B3/B4`). Nævneren Beregning!B4 er 0, fordi Data!B2 og Data!B3 begge er 5.\n\n" +
  "**Vejen:** Beregning!B5 → Resultat!B2 → Resultat!B3.\n\n" +
  "**Rettelse:** `=IF(B4=0,0,B3/B4)`.";

test("Find fejlkilden: en #DIV/0!-kæde over to ark sendes som graf, og svaret peger på kilden med et klikbart link", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.15" }, excel: DIV });
  await select(tp.page, "B3");
  tp.model.reply = () => ({ stream: [ANSWER] });
  await tp.run("Find fejlkilden");
  await expect(tp.page.getByText("Nævneren")).toBeVisible();

  // Modellen får grafen – ikke arket: hver celle med adresse (med arknavn), formel, værdi og fejl, og kilden markeret.
  const sent = prompt(tp);
  expect(sent).toContain("Cell: Resultat!B3");
  expect(sent).toContain("Formula: =B2+1");
  expect(sent).toContain("Precedents of Resultat!B3");
  expect(sent).toContain(
    [
      "L0 Resultat!B3: =B2+1 → #DIV/0! [error]",
      "L1 Resultat!B2: =Beregning!B5*2 → #DIV/0! [error]",
      "L2 Beregning!B5: =B3/B4 → #DIV/0! [error; ORIGIN: the error starts in this cell]",
      "L3 Beregning!B3: 1200",
      "L3 Beregning!B4: =Data!B2-Data!B3 → 0",
      "L4 Data!B2: 5",
      "L4 Data!B3: 5",
    ].join("\n"),
  );
  expect(sent).toContain("A cell marked ORIGIN is where the add-in found that the error starts");
  expect(sent).not.toContain("TRUNCATED:");
  // Grafen er gennemløbet med de direkte præcedenter, niveau for niveau; hele grafen hentes kun, når den er beskåret.
  expect(await traces(tp)).toEqual([
    "getDirectPrecedents Resultat!B3",
    "getDirectPrecedents Resultat!B2",
    "getDirectPrecedents Beregning!B5",
    "getDirectPrecedents Beregning!B4",
  ]);

  // Adresserne i svaret er klikbare; formlen i kodeformat er ikke. Et klik på kilden skifter ark og markerer cellen.
  await expect(tp.page.locator("button.docref")).toHaveText([
    "Beregning!B5",
    "Beregning!B4",
    "Data!B2",
    "Data!B3",
    "Beregning!B5",
    "Resultat!B2",
    "Resultat!B3",
  ]);
  await tp.page.locator("button.docref").first().click();
  await expect
    .poll(async () => (await tp.calls()).filter((c) => c.kind === "activate" || c.kind === "select"))
    .toEqual([
      { kind: "activate", sheet: "Beregning" },
      { kind: "select", address: "Beregning!B5" },
    ]);
  // Sporingen læser kun: intet er skrevet i arkene.
  expect((await tp.calls()).filter((c) => c.kind === "values" || c.kind === "formulas")).toEqual([]);
});

test("Find fejlkilden: brugerens note går med, og en celle uden formel giver en forklaring i stedet for en kørsel", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.15" }, excel: DIV });
  await select(tp.page, "B2");
  tp.model.reply = () => ({ stream: ["Se Beregning!B5."] });
  await tp.page.getByRole("button", { name: "Find fejlkilden", exact: true }).click();
  await tp.page.getByRole("textbox", { name: "Hvad er galt? (valgfrit)" }).fill("Burde være 480");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await expect(tp.page.getByText("Se Beregning!B5.")).toBeVisible();
  expect(prompt(tp)).toContain("The user's note: Burde være 480");
  expect(prompt(tp)).toContain("Precedents of Resultat!B2");

  const before = tp.model.requests.length;
  await select(tp.page, "A2");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await expect(tp.page.getByText("Markér en celle, der indeholder en formel.")).toBeVisible();
  expect(tp.model.requests.length).toBe(before);
});

test("Forklar formel får formlens nærmeste præcedenter med, så forklaringen kan sige, hvor tallene kommer fra", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.15" }, excel: DIV });
  await select(tp.page, "B2");
  tp.model.reply = () => ({ stream: ["Formlen ganger Beregning!B5 med 2."] });
  await tp.run("Forklar formel");
  await expect(tp.page.getByText("Formlen ganger")).toBeVisible();
  const sent = prompt(tp);
  expect(sent).toContain("Also say where the numbers come from, using the traced cells.");
  expect(sent).toContain("The cells the formula is calculated from:");
  expect(sent).toContain("L1 Beregning!B5: =B3/B4 → #DIV/0! [error; ORIGIN: the error starts in this cell]");
  // Kun de nærmeste niveauer: det, kilden bruger, er med, men ikke længere ned.
  expect(sent).toContain("L2 Beregning!B4: =Data!B2-Data!B3 → 0");
  expect(sent).not.toContain("Data!B2: 5");
  await expect(tp.page.locator("button.docref")).toHaveText(["Beregning!B5"]);
});

test("agenten sporer med tracePrecedents og traceDependents i sin læseløkke, før den laver planen", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.15" }, excel: DIV });
  await select(tp.page, "B3");
  let round = 0;
  tp.model.reply = () =>
    ++round === 1
      ? {
          json: JSON.stringify({
            summary: "Jeg sporer fejlen først.",
            actions: [
              { action: "tracePrecedents", address: "B3", reason: "Find kilden" },
              { action: "traceDependents", address: "Data!B2", reason: "Hvad bruger tallet?" },
            ],
          }),
        }
      : {
          json: JSON.stringify({
            summary: "Retter divisionen i Beregning!B5.",
            actions: [{ action: "setFormula", address: "Beregning!B5", formula: "=IF(B4=0,0,B3/B4)", reason: "Undgå division med 0" }],
          }),
        };
  await tp.page.getByRole("button", { name: "Agent: redigér dokumentet", exact: true }).click();
  await tp.page.getByRole("textbox").fill("Ret fejlen i B3");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  // Kun den færdige plan vises; sporingerne er sket uden bekræftelse og uden at ændre noget.
  await expect(tp.page.getByText("Retter divisionen i Beregning!B5.")).toBeVisible();
  await expect(tp.page.getByText("Jeg sporer fejlen først.")).toHaveCount(0);
  expect(round).toBe(2);
  const [first, second] = tp.model.requests.slice(-2).map((r) => r.messages.map((m) => m.content).join("\n"));
  expect(first).toContain("- tracePrecedents {address}");
  expect(first).toContain("- traceDependents {address}");
  expect(first).not.toContain("Precedents of");
  // Anden runde har begge grafer som data: præcedenterne fra det aktive ark og de afhængige på tværs af ark.
  expect(second).toContain("Cells you asked to read");
  expect(second).toContain("Precedents of Resultat!B3");
  expect(second).toContain("L2 Beregning!B5: =B3/B4 → #DIV/0! [error; ORIGIN: the error starts in this cell]");
  expect(second).toContain("Dependents of Data!B2");
  expect(second).toContain(
    ["L1 Beregning!B4: =Data!B2-Data!B3 → 0", "L2 Beregning!B5: =B3/B4 → #DIV/0! [error]", "L3 Resultat!B2: =Beregning!B5*2 → #DIV/0! [error]"].join("\n"),
  );
  expect((await tp.calls()).filter((c) => c.kind === "values" || c.kind === "formulas")).toEqual([]);

  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
  expect((await tp.calls()).filter((c) => c.kind === "formulas")).toContainEqual(
    expect.objectContaining({ address: "Beregning!B5", formulas: [["=IF(B4=0,0,B3/B4)"]] }),
  );
});

/** En kæde på 30 celler (A1 = A2+1, A2 = A3+1, …): dybere end loftet på 25 niveauer. */
const CHAIN = {
  sheet: "Ark1",
  values: Array.from({ length: 30 }, (_, i) => [i === 29 ? 1 : `=A${i + 2}+1`]) as (string | number)[][],
  results: Object.fromEntries(Array.from({ length: 29 }, (_, i) => [`A${i + 1}`, 30 - i])),
};

test("en beskåret graf siger det: med ExcelApi 1.14 fortæller Excels samlede graf, hvad der mangler", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.14" }, excel: CHAIN });
  await select(tp.page, "A1");
  tp.model.reply = () => ({ stream: ["Kæden fortsætter under Ark1!A26."] });
  await tp.run("Find fejlkilden");
  await expect(tp.page.getByText("Kæden fortsætter")).toBeVisible();
  const sent = prompt(tp);
  expect(sent).toContain("L25 Ark1!A26: =A27+1 → 5");
  expect(sent).not.toContain("Ark1!A27:");
  expect(sent).toContain("TRUNCATED: the trace stopped here (limit: 25 levels, 200 cells). Not traced further: Ark1!A26.");
  expect(sent).toContain("According to Excel, the whole graph of precedents covers 29 cells in: Ark1!A2, Ark1!A3");
  expect(sent).toContain("If the list is marked TRUNCATED and the source is not in it, say so");
  const used = await traces(tp);
  expect(used.filter((c) => c.startsWith("getDirectPrecedents"))).toHaveLength(25);
  expect(used.filter((c) => c.startsWith("getPrecedents"))).toEqual(["getPrecedents Ark1!A1"]);
});

test("uden ExcelApi 1.14 bruges kun de direkte præcedenter: sporingen virker, og agenten får ikke traceDependents", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.12" }, excel: DIV });
  await select(tp.page, "B3");
  tp.model.reply = () => ({ stream: [ANSWER] });
  await tp.run("Find fejlkilden");
  await expect(tp.page.getByText("Nævneren")).toBeVisible();
  expect(prompt(tp)).toContain("L2 Beregning!B5: =B3/B4 → #DIV/0! [error; ORIGIN: the error starts in this cell]");
  expect((await traces(tp)).every((c) => c.startsWith("getDirectPrecedents"))).toBe(true);
  await expect(tp.page.locator("button.docref").first()).toHaveText("Beregning!B5");

  // Afhængige kræver ExcelApi 1.13: agenten får kun den sporing, værten kan.
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Intet at ændre.", actions: [] }) });
  await tp.page.getByRole("button", { name: "Alle handlinger" }).click();
  await tp.page.getByRole("button", { name: "Agent: redigér dokumentet", exact: true }).click();
  await tp.page.getByRole("textbox").fill("Ret fejlen i B3");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await expect(tp.page.getByText("Intet at ændre.")).toBeVisible();
  expect(prompt(tp)).toContain("- tracePrecedents {address}");
  expect(prompt(tp)).not.toContain("- traceDependents {address}");
});

test("uden ExcelApi 1.14 siges det stadig, når grafen er beskåret", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.12" }, excel: CHAIN });
  await select(tp.page, "A1");
  tp.model.reply = () => ({ stream: ["Kæden fortsætter under Ark1!A26."] });
  await tp.run("Find fejlkilden");
  await expect(tp.page.getByText("Kæden fortsætter")).toBeVisible();
  expect(prompt(tp)).toContain("TRUNCATED: the trace stopped here (limit: 25 levels, 200 cells). Not traced further: Ark1!A26.");
  expect(prompt(tp)).not.toContain("According to Excel");
  expect((await traces(tp)).some((c) => c.startsWith("getPrecedents"))).toBe(false);
});

test("uden ExcelApi 1.12 forklares det, at Excel ikke kan spore; Forklar formel og agenten virker uden sporing", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: DIV });
  await select(tp.page, "B3");
  const before = tp.model.requests.length;
  await tp.run("Find fejlkilden");
  await expect(tp.page.getByText("Denne Excel kan ikke oplyse, hvilke celler en formel regnes ud fra")).toBeVisible();
  expect(tp.model.requests.length).toBe(before);
  expect(await traces(tp)).toEqual([]);

  tp.model.reply = () => ({ stream: ["Formlen lægger 1 til B2."] });
  await tp.run("Forklar formel");
  await expect(tp.page.getByText("Formlen lægger 1 til B2.")).toBeVisible();
  expect(prompt(tp)).toContain("Formula: =B2+1");
  expect(prompt(tp)).not.toContain("The cells the formula is calculated from");

  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Intet at ændre.", actions: [] }) });
  await tp.page.getByRole("button", { name: "Alle handlinger" }).click();
  await tp.page.getByRole("button", { name: "Agent: redigér dokumentet", exact: true }).click();
  await tp.page.getByRole("textbox").fill("Ret fejlen i B3");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await expect(tp.page.getByText("Intet at ændre.")).toBeVisible();
  expect(prompt(tp)).toContain("readRange {address}");
  expect(prompt(tp)).not.toContain("- tracePrecedents {address}");
  expect(await traces(tp)).toEqual([]);
});

test("formler uden henvisninger, hele kolonner og store områder: Excels ItemNotFound er et svar, og kun den brugte del læses", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.15" },
    excel: {
      sheet: "Ark1",
      values: [
        ["=B1+C1+D1", "=5*2", "=SUM(Tal!B:B)", "=E1", 7],
        ["Sats", 0.25, "", "", ""],
      ],
      results: { A1: 107, B1: 10, C1: 90, D1: 7 },
      sheets: [
        {
          sheet: "Tal",
          values: [["Måned", "Beløb"], ...Array.from({ length: 20 }, (_, i) => [`M${i + 1}`, i === 4 ? "=Ark1!B2*40" : 4])],
          results: { B6: 10 },
        },
      ],
    },
  });
  await select(tp.page, "A1");
  tp.model.reply = () => ({ stream: ["A1 er summen af Ark1!B1, Ark1!C1 og Ark1!D1."] });
  await tp.run("Find fejlkilden");
  await expect(tp.page.getByText("A1 er summen")).toBeVisible();
  const sent = prompt(tp);
  // B1 har ingen præcedenter (Excel svarer ItemNotFound for hele rundturen); C1 og D1 spores alligevel.
  expect(sent).toContain("L1 Ark1!B1: =5*2 → 10\nL1 Ark1!C1: =SUM(Tal!B:B) → 90\nL1 Ark1!D1: =E1 → 7");
  expect(sent).toContain("L2 Ark1!E1: 7");
  // Hele kolonnen læses kun i sin brugte del: formlen og de første værdier vises, resten tælles.
  expect(sent).toContain('L2 Tal!B:B: a range of 21 cells; 12 listed below, not listed (none of them has an error): 9 values.\nL2 Tal!B1: "Beløb"');
  expect(sent).toContain("L2 Tal!B6: =Ark1!B2*40 → 10");
  expect(sent).toContain("L3 Ark1!B2: 0.25");
  expect(sent).not.toContain("could not");
  // Rundturen med B1, C1 og D1 fejlede på B1, så C1 er spurgt igen i en mindre del.
  expect((await traces(tp)).filter((c) => c === "getDirectPrecedents Ark1!C1")).toHaveLength(2);
  expect((await traces(tp)).some((c) => c.startsWith("getPrecedents"))).toBe(false);
});
