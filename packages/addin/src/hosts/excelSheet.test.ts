import { describe, expect, it } from "vitest";
import { buildOverview, colIndex, colName, groups, range, toTsv } from "./excelSheet";

describe("kolonnebogstaver (#57)", () => {
  it("colName: 0-baseret index → bogstaver", () => {
    expect([0, 25, 26, 27, 701, 702, 16383].map(colName)).toEqual(["A", "Z", "AA", "AB", "ZZ", "AAA", "XFD"]);
  });

  it("colIndex: adresse → 0-baseret index, også for et område", () => {
    expect(colIndex("A1")).toBe(0);
    expect(colIndex("C7:E9")).toBe(2);
    expect(colIndex("XFD1")).toBe(16383);
    expect(colIndex("1")).toBe(0);
  });

  it("colName og colIndex er hinandens modsatte", () => {
    for (const i of [0, 1, 25, 26, 51, 52, 701, 702, 16383]) expect(colIndex(`${colName(i)}1`)).toBe(i);
  });
});

describe("toTsv, range og groups (#57)", () => {
  it("toTsv: tabulator og linjeskift i celler bliver mellemrum, tomme celler tom tekst", () => {
    expect(
      toTsv([
        ["a\tb", "c\nd"],
        [null, 3],
      ]),
    ).toBe("a b\tc d\n\t3");
  });

  it("range: fra og med, til og med", () => {
    expect(range(2, 5)).toEqual([2, 3, 4, 5]);
    expect(range(3, 3)).toEqual([3]);
  });

  it("groups: sorterede rækkenumre → sammenhængende stykker", () => {
    expect(groups([1, 2, 3, 7, 9, 10])).toEqual([
      [1, 3],
      [7, 7],
      [9, 10],
    ]);
    expect(groups([])).toEqual([]);
  });
});

describe("buildOverview: Excel-oversigt", () => {
  const values = [
    ["Navn", "Beløb"],
    ["A", 10],
    ["B", 5],
  ];

  it("kolonnebogstaver følger used range (ikke arknavnet)", () => {
    const o = buildOverview(["Sheet1 Sheet1!C3:D5"], "Sheet1", "C3:D5", 3, 2, values);
    expect(o).toContain("Headers (row 1): C=Navn D=Beløb");
    expect(o).toContain("D: 3 non-empty, numeric sum=15");
  });

  it("uden statistik vises de første rækker i stedet", () => {
    const o = buildOverview([], "Ark1", "A1:B3", 3, 2, values, false);
    expect(o).toContain("First rows:");
    expect(o).not.toContain("Column stats");
  });

  it("flere ark nævnes; min og max regnes kun over tal", () => {
    const o = buildOverview(["Ark1 Ark1!A1:B3", "Ark2 Ark2!A1:A2"], "Ark1", "A1:B3", 3, 2, values);
    expect(o).toContain("Active sheet 'Ark1', used range A1:B3 (3 rows x 2 columns)");
    expect(o).toContain("All sheets: Ark1 Ark1!A1:B3 | Ark2 Ark2!A1:A2");
    expect(o).toContain("A: 3 non-empty B: 3 non-empty, numeric sum=15 min=5 max=10");
  });

  it("siger, når statistikken kun dækker en del af used range", () => {
    const o = buildOverview([], "Ark1", "A1:C9000", 9000, 3, values);
    expect(o).toContain("Stats cover the first 3 of 9000 rows.");
    expect(o).toContain("Stats cover the first 2 of 3 columns.");
  });

  it("uden tekst i første række vises ingen overskrifter", () => {
    const o = buildOverview([], "Ark1", "A1:B2", 2, 2, [
      [1, 2],
      [3, 4],
    ]);
    expect(o).not.toContain("Headers");
    expect(o).toContain("A: 2 non-empty, numeric sum=4 min=1 max=3");
  });
});
