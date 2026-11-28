import { describe, expect, it } from "vitest";
import { boxAddress, buildOverview, colIndex, colName, countOverwrite, fillPieces, groups, lastRows, range, toTsv, writeTarget } from "./excelSheet";

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

describe("advarsel før overskrivning (#75)", () => {
  // Som Office.js giver `formulas`: formler som tekst med "=", ellers værdien, og "" for en tom celle.
  const existing = [
    ["Region", "Beløb", ""],
    ["Nord", 120, "=B2*2"],
    ["", 0, "=SUM(C2:C2)"],
  ];

  it("countOverwrite: ryd tæller alle celler med indhold og heraf formlerne", () => {
    expect(countOverwrite(existing)).toEqual({ cells: 7, formulas: 2 });
    expect(countOverwrite([["", null, undefined]])).toEqual({ cells: 0, formulas: 0 });
    expect(countOverwrite([])).toEqual({ cells: 0, formulas: 0 });
  });

  it("countOverwrite: tallet 0, FALSK og en formel, der giver tom tekst, er indhold", () => {
    expect(countOverwrite([[0, false, '=""']])).toEqual({ cells: 3, formulas: 1 });
  });

  it("countOverwrite: null rører ikke cellen, og uændrede værdier tæller ikke med", () => {
    const incoming = [
      ["Region", "Omsætning", "Dobbelt"],
      ["Nord", null, 240],
      ["Syd", 0, null],
    ];
    // B1 får ny tekst og C2 mister sin formel; A1, A2 og B3 er uændrede, B2 og C3 røres ikke, og C1 og A3 var tomme.
    expect(countOverwrite(existing, incoming)).toEqual({ cells: 2, formulas: 1 });
  });

  it("countOverwrite: samme formel igen er ingen overskrivning, en anden er", () => {
    expect(countOverwrite([["=B2*2"]], [["=B2*2"]])).toEqual({ cells: 0, formulas: 0 });
    expect(countOverwrite([["=B2*2"]], [["=B2*3"]])).toEqual({ cells: 1, formulas: 1 });
    expect(countOverwrite([[120]], [["120"]])).toEqual({ cells: 0, formulas: 0 });
  });

  it("countOverwrite: værdier uden for det læste område ignoreres", () => {
    expect(countOverwrite([[1]], [[2, 3], [4]])).toEqual({ cells: 1, formulas: 0 });
  });

  it("writeTarget: setValues skriver fra øverste venstre celle, så langt værdierne rækker; korte rækker fyldes med null", () => {
    expect(writeTarget({ action: "setValues", address: "Salg!$B$2:C9", values: [[1, 2, 3], [4]], reason: "" })).toEqual({
      address: "B2:D3",
      incoming: [
        [1, 2, 3],
        [4, null, null],
      ],
    });
    expect(writeTarget({ action: "setValues", address: "A5", values: [["Total"]], reason: "" })).toEqual({ address: "A5", incoming: [["Total"]] });
  });

  it("writeTarget: setFormula rammer én celle (med = foran), og clear hele området uden nye værdier", () => {
    expect(writeTarget({ action: "setFormula", address: "F2", formula: " SUM(B2:B4)", reason: "" })).toEqual({ address: "F2", incoming: [["=SUM(B2:B4)"]] });
    expect(writeTarget({ action: "clear", address: "'Mit ark'!D4:B2", reason: "" })).toEqual({ address: "B2:D4" });
  });

  it("writeTarget: handlinger, der ikke overskriver celler, og ugyldige adresser giver intet", () => {
    expect(writeTarget({ action: "setNumberFormat", address: "A1:B2", format: "0%", reason: "" })).toBeUndefined();
    expect(writeTarget({ action: "deleteRows", address: "A2:B2", reason: "" })).toBeUndefined();
    expect(writeTarget({ action: "clear", address: "ikke en adresse", reason: "" })).toBeUndefined();
    expect(writeTarget({ action: "setValues", address: "A1", values: [[]], reason: "" })).toBeUndefined();
  });
});

describe("markering af ændrede celler (#75)", () => {
  it("boxAddress og lastRows", () => {
    expect(boxAddress({ r0: 1, c0: 1, r1: 3, c1: 27 })).toBe("B2:AB4");
    expect(boxAddress({ r0: 0, c0: 0, r1: 0, c1: 0 })).toBe("A1");
    expect(lastRows("A2:C6", 2)).toBe("A5:C6");
    expect(lastRows("A2:C6", 1)).toBe("A6:C6");
    // Flere rækker, end området har, giver hele området.
    expect(lastRows("A2:C3", 9)).toBe("A2:C3");
  });

  it("fillPieces: et område helt uden fyld ryddes i ét stykke", () => {
    expect(
      fillPieces("B2:C3", [
        [undefined, undefined],
        [undefined, undefined],
      ]),
    ).toEqual([{ address: "B2:C3" }]);
  });

  it("fillPieces: et område, hvor alle celler har fyld, får farverne tilbage i ét stykke", () => {
    const fills = [
      ["#FFFF00", "#FFFF00"],
      ["#00FF00", "#FFFF00"],
    ];
    expect(fillPieces("B2:C3", fills)).toEqual([{ address: "B2:C3", colors: fills }]);
  });

  it("fillPieces: blandet fyld deles i sammenhængende stykker pr. række, og hver celle er med præcis én gang", () => {
    const pieces = fillPieces("B2:E3", [
      ["#FFFF00", "#00FF00", undefined, undefined],
      [undefined, "#FFFF00", undefined, "#FF0000"],
    ]);
    expect(pieces).toEqual([
      { address: "B2:C2", colors: [["#FFFF00", "#00FF00"]] },
      { address: "D2:E2" },
      { address: "B3" },
      { address: "C3", colors: [["#FFFF00"]] },
      { address: "D3" },
      { address: "E3", colors: [["#FF0000"]] },
    ]);
  });
});
