import { expect, test } from "./fixtures/test";

/** Agent-tilstand i Excel og PowerPoint: formateringen bevares, og fortryd gendanner også formatering. */

type Sheet = {
  values: unknown[][];
  formats: Record<string, { numberFormat?: string; bold?: boolean; fill?: string }>;
  tables: { name: string; range: string }[];
};
const sheet = (page: import("@playwright/test").Page) => page.evaluate(() => (window as unknown as { __miai: { sheet: () => Sheet } }).__miai.sheet());

test("Excel: ny række i tabellen og en total får tabellens og kolonnernes format; fortryd gendanner værdier og formater", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Region", "Beløb", "Andel"],
        ["Jylland", 120, 0.4],
        ["Sjælland", 200, 0.6],
      ],
      formats: { "A1:C1": { bold: true }, "B2:B3": { numberFormat: "#,##0 kr" }, "C2:C3": { numberFormat: "0%" } },
      tables: [{ name: "Salg", range: "A1:C3" }],
    },
  });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Tilføjer Fyn og en total.",
      actions: [
        { action: "appendRows", table: "Salg", values: [["Fyn", 70, null]], reason: "Ny region" },
        { action: "setValues", address: "A5", values: [["Total"]], reason: "Etiket" },
        { action: "setFormula", address: "B5", formula: "=SUM(B2:B4)", reason: "Sum" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();
  // Modellen ser tabellen og talformaterne.
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain("- Salg (Ark1!A1:C3): Region | Beløb | Andel");
  expect(prompt).toContain("Number formats: B: #,##0 kr; C: 0%");
  await expect(tp.page.getByRole("checkbox", { name: "Tilføj 1 rækker i tabellen Salg" })).toBeChecked();

  await tp.page.getByRole("button", { name: "Udfør 3 handlinger" }).click();
  await expect(tp.page.getByText("3 af 3 handlinger udført.")).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.values.slice(3)).toEqual([
    ["Fyn", 70, ""],
    ["Total", "=SUM(B2:B4)", ""],
  ]);
  expect(after.tables).toEqual([{ name: "Salg", range: "A1:C4" }]);
  // Tabelrækken har tabellens formater; totalen har kolonnens talformat (ikke "General").
  expect(after.formats.B4).toMatchObject({ numberFormat: "#,##0 kr" });
  expect(after.formats.C4).toMatchObject({ numberFormat: "0%" });
  expect(after.formats.B5).toMatchObject({ numberFormat: "#,##0 kr" });
  // Overskrifterne er urørt.
  expect(after.formats.A1).toMatchObject({ bold: true });

  await expect(tp.page.getByText(/også deres formatering/)).toBeVisible();
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  const undone = await sheet(tp.page);
  expect(undone.values.filter((r) => r.some((v) => v !== ""))).toEqual(before.values);
  expect(undone.tables).toEqual(before.tables);
  expect(undone.formats).toEqual(before.formats);
});

type Deck = {
  id: string;
  layout: string;
  shapes: { name: string; text: string; style: Record<string, unknown>; marks: { start: number; end: number; bold?: boolean }[] }[];
}[];
const deck = (page: import("@playwright/test").Page) => page.evaluate(() => (window as unknown as { __miai: { deck: () => Deck } }).__miai.deck());
const withoutIds = (d: Deck) => d.map(({ id: _, ...rest }) => rest);

const SLIDES = [
  {
    layout: "Title Slide",
    shapes: [
      { name: "Title 1", ph: "Title", text: "Kvartalsrapport" },
      { name: "Subtitle 2", ph: "Subtitle", text: "Q3 2026" },
    ],
  },
  {
    layout: "Title and Content",
    shapes: [
      { name: "Title 1", ph: "Title", text: "Resultater" },
      {
        name: "Content 2",
        ph: "Body",
        text: "Omsætning +12 % i Q3\nMargin 38 %",
        style: { color: "#C00000", size: 24 },
        marks: [{ start: 0, end: 9, bold: true }],
      },
    ],
  },
  {
    layout: "Title and Content",
    shapes: [
      { name: "Title 1", ph: "Title", text: "Gammelt slide" },
      { name: "Content 2", ph: "Body", text: "Slet mig" },
    ],
  },
];
const LAYOUTS = [
  { name: "Title Slide", placeholders: ["Title", "Subtitle"] },
  { name: "Title and Content", placeholders: ["Title", "Body"] },
];

test("PowerPoint: kun det ændrede ændres, nye slides følger skabelonen, og fortryd lægger slides tilbage med formatering", async ({ open }) => {
  const tp = await open({ host: "PowerPoint", apiSets: { PowerPointApi: "1.8" }, slides: SLIDES, layouts: LAYOUTS });
  const original = await deck(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Opdaterer tallet, tilføjer næste skridt og rydder op.",
      actions: [
        { action: "replace", find: "+12 % i Q3", replace: "+14 % i Q3", reason: "Nyt tal" },
        { action: "setText", shape: "1.2", text: "Q3 2026 – **foreløbig**", reason: "Status" },
        { action: "addSlide", after: 2, layout: "Title and Content", title: "Næste skridt", body: "- Godkend budget\n- Ansæt to", reason: "Opfølgning" },
        { action: "deleteSlide", slide: 3, reason: "Forældet" },
        { action: "moveSlide", slide: 2, to: 1, reason: "Resultater først" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain(
    "Slide 2 (layout: Title and Content)\n  2.1 [title] Title 1: Resultater\n  2.2 [body] Content 2: Omsætning +12 % i Q3 ⏎ Margin 38 %",
  );
  expect(prompt).toContain("Layouts: Title Slide | Title and Content");
  await expect(tp.page.getByRole("checkbox", { name: "Ny tekst i Subtitle 2 (slide 1): “Q3 2026 – foreløbig”" })).toBeChecked();

  await tp.page.getByRole("button", { name: "Udfør 5 handlinger" }).click();
  await expect(tp.page.getByText("5 af 5 handlinger udført.")).toBeVisible();
  const d = await deck(tp.page);
  expect(d.map((s) => s.shapes[0].text)).toEqual(["Resultater", "Kvartalsrapport", "Næste skridt"]);
  // Erstat: kun tallet er skiftet; "Omsætning" er stadig fed, og figurens skabelonformat er urørt.
  expect(d[0].shapes[1]).toMatchObject({
    text: "Omsætning +14 % i Q3\nMargin 38 %",
    style: { color: "#C00000", size: 24 },
    marks: [{ start: 0, end: 9, bold: true }],
  });
  // Ny tekst: fed på det markerede ord.
  expect(d[1].shapes[1]).toMatchObject({ text: "Q3 2026 – foreløbig", marks: [{ start: 10, end: 19, bold: true }] });
  // Nyt slide fra layoutet: teksten står i layoutets pladsholdere (ingen egne tekstbokse), punkttegn fra layoutet.
  expect(d[2]).toMatchObject({ layout: "Title and Content" });
  expect(d[2].shapes.map((s) => [s.text, s.style])).toEqual([
    ["Næste skridt", { fromLayout: "Title and Content" }],
    ["Godkend budget\nAnsæt to", { fromLayout: "Title and Content" }],
  ]);

  await expect(tp.page.getByText(/lægger de berørte slides tilbage med deres formatering/)).toBeVisible();
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(withoutIds(await deck(tp.page))).toEqual(withoutIds(original));
});

test("PowerPoint før 1.8: slides slettes ikke (kan ikke fortrydes), og fortryd siger, at kun teksten gendannes", async ({ open }) => {
  const tp = await open({ host: "PowerPoint", apiSets: { PowerPointApi: "1.4" }, slides: SLIDES, layouts: LAYOUTS });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "",
      actions: [
        { action: "replace", find: "Q3 2026", replace: "Q4 2026", reason: "" },
        { action: "deleteSlide", slide: 3, reason: "" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("1 af 2 handlinger udført.")).toBeVisible();
  await expect(tp.page.getByText(/Sletning af slides kræver en nyere PowerPoint/)).toBeVisible();
  expect((await deck(tp.page)).length).toBe(3);
  await expect(tp.page.getByText(/kan kun gendanne teksten/)).toBeVisible();
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect((await deck(tp.page))[0].shapes[1].text).toBe("Q3 2026");
});
