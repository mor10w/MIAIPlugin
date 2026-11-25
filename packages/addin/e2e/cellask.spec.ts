import type { Page } from "@playwright/test";
import { expect, test } from "./fixtures/test";

/** Regnearksfunktionerne =MIAI.ASK() og =MIAI.MODELS() (#80): kaldt, som Excel ville gøre det for en celle. */

const SHEET = {
  sheet: "Ark1",
  values: [
    ["By", "Land"],
    ["Aarhus", ""],
  ],
};
const EXCEL = { host: "Excel" as const, apiSets: { ExcelApi: "1.10" }, excel: SHEET, customFunctions: true };

type CellResult = { value?: string | number; error?: string; message?: string };
const call = (page: Page, id: string, ...args: unknown[]) =>
  page.evaluate(
    ([fn, a]) => (window as unknown as { __miai: { fn: (id: string, ...a: unknown[]) => Promise<CellResult> } }).__miai.fn(fn as string, ...(a as unknown[])),
    [id, args] as const,
  );
const ask = (page: Page, prompt: unknown, input?: unknown, optimized: boolean | null = null, model: string | null = null) =>
  call(page, "ASK", prompt, input ?? null, optimized, model, {});

for (const mode of ["direct", "bridge"] as const) {
  test(`=MIAI.ASK() svarer i cellen, bruger cachen og sender den rigtige profil (${mode})`, async ({ open }) => {
    const tp = await open({ ...EXCEL, mode });
    tp.model.reply = () => ({ stream: ["Dan", "mark"] });

    expect(await ask(tp.page, "Hvilket land ligger byen i?", [["Aarhus"]])).toEqual({ value: "Danmark" });
    const first = tp.model.requests[0];
    expect(first).toMatchObject({ temperature: 0, seed: 1 });
    expect(first.messages[0].content).toContain("ONE spreadsheet cell");
    expect(first.messages[1].content).toBe("Hvilket land ligger byen i?\n\n<document>\nAarhus\n</document>");

    // Samme celle igen (F9): intet nyt modelkald.
    expect(await ask(tp.page, "Hvilket land ligger byen i?", [["Aarhus"]])).toEqual({ value: "Danmark" });
    expect(tp.model.requests).toHaveLength(1);

    // optimized = SAND: kort svar uden tænkning – og et nyt kald, fordi profilen er en anden.
    tp.model.reply = () => ({ stream: ["1234,5"] });
    expect(await ask(tp.page, "Træk beløbet ud", [["Betalt 1.234,50 kr."]], true)).toEqual({ value: 1234.5 });
    expect(tp.model.requests[1]).toMatchObject({ max_tokens: 256, think: false });
    expect(tp.model.requests[1].messages[0].content).not.toContain("You are MIAI");
  });
}

test("model som argument: bruges, når den findes; et ukendt navn giver #VALUE! uden modelkald", async ({ open }) => {
  const tp = await open({ ...EXCEL, models: ["qwen3:4b", "qwen3:32b"] });
  tp.model.reply = () => ({ stream: ["ok"] });
  expect(await ask(tp.page, "p", [["x"]], null, "qwen3:32b")).toEqual({ value: "ok" });
  expect(tp.model.lastRequest?.model).toBe("qwen3:32b");
  expect(await ask(tp.page, "p", [["x"]], null, "findes-ikke")).toEqual({ error: "#VALUE!", message: "Modellen findes ikke på endpointen" });
  expect(tp.model.requests).toHaveLength(1);
  expect(await call(tp.page, "MODELS")).toEqual({ value: [["qwen3:4b"], ["qwen3:32b"]] });
});

test("fejl vises som Excel-fejl med forklaring: prompten mangler, og modellen kan ikke nås", async ({ open }) => {
  const tp = await open(EXCEL);
  expect(await ask(tp.page, "", [["x"]])).toEqual({ error: "#VALUE!", message: "Prompten mangler" });
  await tp.model.stop();
  expect(await ask(tp.page, "p", [["x"]])).toEqual({ error: "#N/A", message: "MIAI kan ikke nå modellen" });
});

test("mange celler: loftet over samtidige kald, status i task pane, og Stop tømmer køen", async ({ open }) => {
  const tp = await open({ ...EXCEL, settings: { cellConcurrency: 2 } });
  // Et langsomt svar, så køen kan ses.
  tp.model.reply = () => ({ stream: Array.from({ length: 400 }, () => "x") });
  const cells = tp.page.evaluate(() => {
    const m = (window as unknown as { __miai: { fn: (id: string, ...a: unknown[]) => Promise<CellResult> } }).__miai;
    return Promise.all([1, 2, 3, 4, 5, 6].map((i) => m.fn("ASK", "Opsummér", [[`række ${i}`]], null, null, {})));
  });
  const bar = tp.page.getByTestId("cell-bar");
  await expect(bar).toContainText("MIAI.ASK: 2 i gang, 4 i kø");
  await bar.getByRole("button", { name: "Stop MIAI-funktioner" }).click();
  for (const r of await cells) expect(r).toEqual({ error: "#N/A", message: "Stoppet" });
  await expect(bar).toBeHidden();
  expect(tp.model.requests).toHaveLength(2); // de fire i køen nåede aldrig modellen
});

test("Indstillinger → Excel-funktioner: cachen kan ryddes, og modelvælgeren indsættes som dropdown", async ({ open }) => {
  const tp = await open({ ...EXCEL, models: ["qwen3:4b", "qwen3:32b"] });
  tp.model.reply = () => ({ stream: ["Danmark"] });
  await ask(tp.page, "Land?", [["Aarhus"]]);

  await tp.tab("Indstillinger");
  await tp.section("Excel-funktioner");
  await expect(tp.page.getByText("1 svar i cachen.")).toBeVisible();
  await tp.page.getByRole("button", { name: "Ryd cache" }).click();
  await expect(tp.page.getByText("Cachen er ryddet.")).toBeVisible();
  await expect(tp.page.getByText("0 svar i cachen.")).toBeVisible();
  await ask(tp.page, "Land?", [["Aarhus"]]);
  expect(tp.model.requests).toHaveLength(2); // cachen var tom, så modellen blev spurgt igen

  await tp.page.getByRole("button", { name: "Indsæt modelvælger" }).click();
  await expect(tp.page.getByText(/Modelvælgeren er indsat i /)).toBeVisible();
  const calls = await tp.calls();
  expect(calls).toContainEqual({ kind: "sheetAdd", name: "MIAI_modeller" });
  expect(calls).toContainEqual({ kind: "sheetVisibility", sheet: "MIAI_modeller", visibility: "Hidden" });
  expect(calls.find((c) => c.kind === "formulas")).toMatchObject({ formulas: [["=MIAI.MODELS()"]] });
  expect(calls.find((c) => c.kind === "validation")).toMatchObject({ rule: { list: { inCellDropDown: true, source: "=MIAI_modeller!$A$1#" } } });
});

test("audit-loggen i direkte tilstand samler en kørsels cellekald i én post", async ({ open }) => {
  const tp = await open({ ...EXCEL, settings: { auditLog: true } });
  tp.model.reply = (body) => ({ stream: [`svar ${body.messages[1].content.length}`] });
  for (const i of [1, 22, 333]) await ask(tp.page, "p", [[`række ${i}`]], true);
  const log = await tp.page.evaluate(() => JSON.parse(localStorage.getItem("miai.audit.v1") ?? "[]") as { kind: string; task?: string; count?: number }[]);
  expect(log.filter((e) => e.task === "cell-ask")).toMatchObject([{ kind: "task", count: 3 }]);
});

test("uden regnearksfunktioner i manifestet (som i de udgivne manifester) vises intet om dem, og intet registreres", async ({ open }) => {
  // Som rigtig Excel: CustomFunctions findes i siden, selv om manifestet ikke har funktionerne med.
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SHEET });
  expect(await tp.page.evaluate(() => typeof (window as unknown as { CustomFunctions?: { associate?: unknown } }).CustomFunctions?.associate)).toBe("function");
  expect(await tp.page.evaluate(() => (window as unknown as { __miai: { functions: () => string[] } }).__miai.functions())).toEqual([]);
  await tp.tab("Indstillinger");
  await expect(tp.page.getByRole("button", { name: "Generelt", exact: true })).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Excel-funktioner", exact: true })).toHaveCount(0);
});

test("med funktionerne i manifestet registreres ASK og MODELS", async ({ open }) => {
  const tp = await open(EXCEL);
  expect(await tp.page.evaluate(() => (window as unknown as { __miai: { functions: () => string[] } }).__miai.functions().sort())).toEqual(["ASK", "MODELS"]);
});

for (const mode of ["direct", "bridge"] as const) {
  test(`funktionen virker i Excels skjulte udgave af siden, hvor panelet aldrig starter (${mode})`, async ({ open }) => {
    // Regression: funktionen ventede på panelets opstart, som kun sker, når Office melder Excel som
    // vært. I en side, der kun kører funktioner, stod cellen derfor på #BUSY! uden noget kald.
    const tp = await open({ ...EXCEL, mode, hiddenRuntime: true });
    await expect(tp.page.getByText("Denne Office-app understøttes ikke endnu.")).toBeVisible();
    tp.model.reply = () => ({ stream: ["abc"] });
    expect(await ask(tp.page, "Fjern _", [["a_b_c"]], true)).toEqual({ value: "abc" });
    expect(tp.model.requests).toHaveLength(1);
  });
}
