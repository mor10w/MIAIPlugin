import { expect, test } from "./fixtures/test";

/** Excel: oversættelse celle for celle, dataanalyse med arkstatistik, formler. */

const SHEET = {
  sheet: "Sheet1",
  values: [
    ["Status", "Beløb"],
    ["Færdig", 10],
    ["Færdig", 5],
  ] as (string | number)[][],
};

test("oversæt et område: TSV-forhåndsvisning, skrives celle for celle, ingen presence_penalty", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SHEET, endpoint: { presencePenalty: 1.5 } });
  tp.model.reply = () => ({ json: "Status\tAmount\nDone\t10\nDone\t5" });
  await tp.run("Oversæt");

  // TSV hører hjemme i cellerne: "Skriv i celler" er primær, og "Erstat" (én celle) tilbydes ikke.
  await expect(tp.page.getByRole("button", { name: "Skriv i celler" })).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Erstat" })).toHaveCount(0);
  await expect(tp.page.locator(".fui-Card table tr")).toHaveCount(3);

  await tp.page.getByRole("button", { name: "Skriv i celler" }).click();
  await expect(tp.page.getByText("Indsat.")).toBeVisible();
  expect(await tp.calls()).toContainEqual(expect.objectContaining({ kind: "values", values: [["Status", "Amount"], ["Done", "10"], ["Done", "5"]] }));

  // presence_penalty skader tabeller (gentagne værdier) og sendes derfor kun til fritekst.
  expect(tp.model.lastRequest).not.toHaveProperty("presence_penalty");
});

test("analyser data: arkstatistik med korrekte kolonnebogstaver", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SHEET });
  tp.model.reply = () => ({ stream: ["Kolonne B summer til 15."] });
  await tp.run("Analyser data");
  await expect(tp.page.getByText("Kolonne B summer til 15.")).toBeVisible();

  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  // Regression: bogstaverne blev læst fra arknavnet ("Sheet1" → kolonne S).
  expect(prompt).toContain("Headers (row 1): A=Status B=Beløb");
  expect(prompt).toContain("B: 3 non-empty, numeric sum=15");
});

test("generér formel: skrives i cellen og kan fortrydes", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SHEET });
  tp.model.reply = () => ({ json: JSON.stringify({ formula: "=SUM(B2:B3)", explanation: "Summerer beløbene." }) });
  await tp.page.getByRole("button", { name: "Lav formel ud fra beskrivelse", exact: true }).click();
  await tp.page.getByRole("textbox").fill("summen af beløb");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await tp.page.getByRole("button", { name: "Indsæt formel" }).click();
  await tp.page.getByRole("button", { name: "Fortryd" }).click();
  await expect(tp.page.getByText("Fortrudt.")).toBeVisible();

  const formulas = (await tp.calls()).filter((c) => c.kind === "formulas");
  expect(formulas[0]).toMatchObject({ formulas: [["=SUM(B2:B3)"]] });
  expect(formulas[1]).toMatchObject({ formulas: [["Status"]] }); // tidligere indhold gendannet
});
