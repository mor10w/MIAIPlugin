import { rangeBox, type TraceKind } from "@miai/shared";
import { describe, expect, it } from "vitest";
import { colName } from "./excelSheet";
import {
  AREA_SAMPLE,
  areaBox,
  MAX_AREA_READ,
  parseAreas,
  refText,
  renderTrace,
  startCell,
  TRACE_LIMITS,
  traceGraph,
  type AreaCells,
  type CellRef,
  type TraceArea,
  type TraceCell,
  type TraceReader,
} from "./excelTrace";

/**
 * En lille projektmappe i hukommelsen: ark → celle → fast værdi, eller [formel, resultat].
 * Afhængighederne udledes af formlernes henvisninger, og kaldene tælles.
 */
type Cell = string | number | [formula: string, value: string | number];
type Book = Record<string, Record<string, Cell>>;

const REF = /(?:('(?:[^']|'')+'|[\p{L}\p{N}_.]+)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})/gu;

function fakeReader(book: Book, kind: TraceKind = "precedents", opts: { whole?: boolean; fail?: string[]; types?: boolean } = {}) {
  const calls = { direct: 0, read: 0, whole: 0, asked: [] as string[] };
  const sheetOf = (name: string) => Object.keys(book).find((s) => s.toLowerCase() === name.toLowerCase());
  const refs = (sheet: string, formula: string): TraceArea[] =>
    [...formula.slice(1).matchAll(REF)].map((m) => ({
      sheet: m[1] ? m[1].replace(/^'(.*)'$/, "$1").replace(/''/g, "'") : sheet,
      range: m[2].replace(/\$/g, ""),
    }));
  const formulaOf = (c: Cell | undefined) => (Array.isArray(c) ? c[0] : undefined);
  /** Områdets hjørner; en hel kolonne rækker over arkets brugte rækker. */
  const boxOf = (a: TraceArea) => {
    const b = areaBox(a.range);
    if (b) return b;
    const col = rangeBox(`${a.range.split(":")[0]}1`).c0;
    const rows = Object.keys(book[sheetOf(a.sheet)!] ?? {})
      .map((k) => rangeBox(k))
      .filter((x) => x.c0 === col)
      .map((x) => x.r0);
    return rows.length ? { r0: Math.min(...rows), r1: Math.max(...rows), c0: col, c1: col } : undefined;
  };
  const inside = (a: TraceArea, cell: CellRef) => {
    const b = boxOf(a);
    const c = rangeBox(cell.address);
    return !!b && a.sheet.toLowerCase() === cell.sheet.toLowerCase() && c.r0 >= b.r0 && c.r0 <= b.r1 && c.c0 >= b.c0 && c.c0 <= b.c1;
  };
  const precedents = (cell: CellRef) => {
    const f = formulaOf(book[sheetOf(cell.sheet)!]?.[cell.address]);
    return f ? refs(sheetOf(cell.sheet)!, f) : [];
  };
  const dependents = (cell: CellRef): TraceArea[] =>
    Object.entries(book).flatMap(([sheet, cells]) =>
      Object.entries(cells).flatMap(([address, c]) => {
        const f = formulaOf(c);
        return f && refs(sheet, f).some((a) => inside(a, cell)) ? [{ sheet, range: address }] : [];
      }),
    );
  const reader: TraceReader = {
    async direct(cells) {
      calls.direct++;
      calls.asked.push(...cells.map(refText));
      return cells.map((c) => (opts.fail?.includes(refText(c)) ? undefined : kind === "precedents" ? precedents(c) : dependents(c)));
    },
    async read(areas) {
      calls.read++;
      return areas.map((a): AreaCells | undefined => {
        const name = sheetOf(a.sheet);
        if (!name) return undefined;
        const b = boxOf(a);
        if (!b) return { row: 0, col: 0, formulas: [], values: [] };
        const total = (b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1);
        const rows = Math.min(b.r1 - b.r0 + 1, Math.max(1, Math.floor(MAX_AREA_READ / (b.c1 - b.c0 + 1))));
        const grid = (pick: (c: Cell | undefined) => unknown) =>
          Array.from({ length: rows }, (_, r) => Array.from({ length: b.c1 - b.c0 + 1 }, (_, c) => pick(book[name][`${colName(b.c0 + c)}${b.r0 + r + 1}`])));
        const values = grid((c) => (Array.isArray(c) ? c[1] : (c ?? "")));
        return {
          row: b.r0,
          col: b.c0,
          formulas: grid((c) => (Array.isArray(c) ? c[0] : (c ?? ""))),
          values,
          ...(opts.types
            ? { types: values.map((row) => row.map((v) => (typeof v === "string" && v.startsWith("#") ? "Error" : v === "" ? "Empty" : "String"))) }
            : {}),
          ...(rows * (b.c1 - b.c0 + 1) < total ? { total } : {}),
        };
      });
    },
    ...(opts.whole
      ? {
          async whole(cell: CellRef) {
            calls.whole++;
            // Hele grafen: alle områder, der nås fra cellen.
            const out: TraceArea[] = [];
            const todo = [cell];
            const seen = new Set<string>();
            while (todo.length) {
              const c = todo.pop()!;
              for (const a of kind === "precedents" ? precedents(c) : dependents(c)) {
                const key = `${a.sheet}!${a.range}`;
                if (seen.has(key)) continue;
                seen.add(key);
                out.push(a);
                const b = boxOf(a);
                if (b)
                  for (let r = b.r0; r <= b.r1; r++) for (let col = b.c0; col <= b.c1; col++) todo.push({ sheet: a.sheet, address: `${colName(col)}${r + 1}` });
              }
            }
            return out;
          },
        }
      : {}),
  };
  return { reader, calls };
}

const cells = (g: { items: { type: string }[] }) => g.items.filter((x) => x.type === "cell") as unknown as TraceCell[];
const names = (g: { items: { type: string }[] }) => cells(g).map((c) => `L${c.depth} ${refText(c)}`);

/** #DIV/0! over to ark: Resultat!B2 ← Beregning!B5 (=B3/B4) ← B4 (=Data!B2-Data!B3 = 0). */
const DIV: Book = {
  Resultat: { A2: "Pris pr. stk.", B2: ["=Beregning!B5*2", "#DIV/0!"], B3: ["=B2+1", "#DIV/0!"] },
  Beregning: { B3: 1200, B4: ["=Data!B2-Data!B3", 0], B5: ["=B3/B4", "#DIV/0!"] },
  Data: { B2: 5, B3: 5 },
};

describe("adresser fra Excel (#76)", () => {
  it("parseAreas: flere områder i én tekst, arknavne i apostroffer, $ og semikolon", () => {
    expect(parseAreas(["Data!$B$2:$B$13, Data!D4", "'Mit, ark'!a1;'Bob''s'!C3:C4"], "Ark1")).toEqual([
      { sheet: "Data", range: "B2:B13" },
      { sheet: "Data", range: "D4" },
      { sheet: "Mit, ark", range: "A1" },
      { sheet: "Bob's", range: "C3:C4" },
    ]);
    // Uden arknavn gælder cellens eget ark; tomme stykker springes over.
    expect(parseAreas(["B2:B3, ", ""], "Ark1")).toEqual([{ sheet: "Ark1", range: "B2:B3" }]);
  });

  it("areaBox: celler og områder, men ikke hele kolonner eller rækker", () => {
    expect(areaBox("B2:D4")).toEqual({ r0: 1, c0: 1, r1: 3, c1: 3 });
    expect(areaBox("AA10")).toEqual({ r0: 9, c0: 26, r1: 9, c1: 26 });
    expect(areaBox("B:B")).toBeUndefined();
    expect(areaBox("2:5")).toBeUndefined();
  });

  it("startCell: første celle i adressen, med arknavn fra adressen eller det aktive ark", () => {
    expect(startCell("Ark1!$B$2:C5", "Andet")).toEqual({ sheet: "Ark1", address: "B2" });
    expect(startCell("b7", "Ark1")).toEqual({ sheet: "Ark1", address: "B7" });
    expect(startCell("'Mit ark'!C3", "Ark1")).toEqual({ sheet: "Mit ark", address: "C3" });
  });

  it("refText: arknavne med mellemrum eller tegn står i apostroffer, som Excel skriver dem", () => {
    expect(refText({ sheet: "Data", address: "B4" })).toBe("Data!B4");
    expect(refText({ sheet: "Mit ark", address: "B4" })).toBe("'Mit ark'!B4");
    expect(refText({ sheet: "Bob's", address: "B4" })).toBe("'Bob''s'!B4");
    expect(refText({ sheet: "Salg_2026.v2", address: "B4" })).toBe("Salg_2026.v2!B4");
  });
});

describe("gennemløb af præcedenter (#76)", () => {
  it("en #DIV/0!-kæde over to ark: kun fejlceller følges, kilden markeres, og det, kildens formel bruger, kommer med", async () => {
    const { reader, calls } = fakeReader(DIV);
    const g = await traceGraph({ sheet: "Resultat", address: "B3" }, "precedents", reader);
    expect(names(g)).toEqual([
      "L0 Resultat!B3",
      "L1 Resultat!B2",
      "L2 Beregning!B5",
      "L3 Beregning!B3",
      "L3 Beregning!B4",
      // Nævneren er 0, fordi Data!B2 − Data!B3 er 0: ét niveau under kilden, og ikke længere.
      "L4 Data!B2",
      "L4 Data!B3",
    ]);
    expect(
      cells(g)
        .filter((c) => c.origin)
        .map(refText),
    ).toEqual(["Beregning!B5"]);
    expect(
      cells(g)
        .filter((c) => c.error)
        .map(refText),
    ).toEqual(["Resultat!B3", "Resultat!B2", "Beregning!B5"]);
    expect(g).toMatchObject({ followedErrors: true, omitted: 0, notFollowed: [], failed: [] });
    // Én forespørgsel pr. niveau; de faste værdier på sidste niveau spørges der ikke om.
    expect(calls.direct).toBe(4);
    expect(calls.asked).toEqual(["Resultat!B3", "Resultat!B2", "Beregning!B5", "Beregning!B4"]);

    const text = renderTrace(g);
    expect(text).toContain("Precedents of Resultat!B3");
    expect(text).toContain("L0 Resultat!B3: =B2+1 → #DIV/0! [error]");
    expect(text).toContain("L2 Beregning!B5: =B3/B4 → #DIV/0! [error; ORIGIN: the error starts in this cell]");
    expect(text).toContain("L3 Beregning!B3: 1200\nL3 Beregning!B4: =Data!B2-Data!B3 → 0\nL4 Data!B2: 5");
    expect(text).toContain("Only cells with an error value were traced further");
    expect(text).not.toContain("TRUNCATED");
  });

  it("uden fejl i startcellen følges alle formler, og tekst, tal og tomme celler vises forskelligt", async () => {
    const { reader } = fakeReader({
      Ark1: { A1: ["=B1+C1", 15], B1: ["=D1*2", 10], C1: 5, D1: 5, E1: ["=A1&F1&G1", "15kr"], F1: "kr" },
    });
    const g = await traceGraph({ sheet: "Ark1", address: "E1" }, "precedents", reader);
    expect(names(g)).toEqual(["L0 Ark1!E1", "L1 Ark1!A1", "L1 Ark1!F1", "L1 Ark1!G1", "L2 Ark1!B1", "L2 Ark1!C1", "L3 Ark1!D1"]);
    expect(g.followedErrors).toBe(false);
    const text = renderTrace(g);
    expect(text).toContain('L0 Ark1!E1: =A1&F1&G1 → "15kr"');
    expect(text).toContain('L1 Ark1!F1: "kr"');
    expect(text).toContain("L1 Ark1!G1: (empty)");
    expect(text).not.toContain("ORIGIN");
  });

  it("en tom nævner vises, og formlen, der dividerer med den, er kilden", async () => {
    const { reader } = fakeReader({ Ark1: { A1: 10, C1: ["=A1/B1", "#DIV/0!"] } });
    const g = await traceGraph({ sheet: "Ark1", address: "C1" }, "precedents", reader);
    const text = renderTrace(g);
    expect(text).toContain("L0 Ark1!C1: =A1/B1 → #DIV/0! [error; ORIGIN: the error starts in this cell]");
    expect(text).toContain("L1 Ark1!B1: (empty)");
  });

  it("fejlværdier: en formel uden henvisninger og en fast fejlværdi er selv kilden; Excels værdityper har forrang for teksten", async () => {
    const own = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ Ark1: { A1: ["=1/0", "#DIV/0!"] } }).reader);
    expect(cells(own)[0]).toMatchObject({ error: "#DIV/0!", origin: true });
    expect(renderTrace(own)).toContain("The cell's formula uses no other cells.");

    const typed = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ Ark1: { A1: ["=B1*2", "#N/A"], B1: "#N/A" } }).reader);
    expect(cells(typed).map((c) => [refText(c), c.origin ?? false])).toEqual([
      ["Ark1!A1", false],
      ["Ark1!B1", true],
    ]);

    // Alle Excels fejlværdier genkendes på teksten, når værten ikke oplyser værdityper …
    for (const e of ["#REF!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"]) {
      const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ Ark1: { A1: ["=B1", e], B1: 1 } }).reader);
      expect(cells(g)[0].error, e).toBe(e);
    }
    // … og en tekst, der blot begynder med #, er ikke en fejl.
    const text = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ Ark1: { A1: ["=B1", "#hashtag"], B1: "#hashtag" } }).reader);
    expect(cells(text).some((c) => c.error)).toBe(false);
    const withTypes = await traceGraph(
      { sheet: "Ark1", address: "A1" },
      "precedents",
      fakeReader({ Ark1: { A1: ["=B1", "#REF!"], B1: 1 } }, "precedents", { types: true }).reader,
    );
    expect(cells(withTypes)[0].error).toBe("#REF!");
  });

  it("cirkulære henvisninger: hver celle ses én gang, og gennemløbet stopper", async () => {
    const { reader, calls } = fakeReader({ Ark1: { A1: ["=B1+1", 0], B1: ["=C1+1", 0], C1: ["=A1+1", 0] } });
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader);
    expect(names(g)).toEqual(["L0 Ark1!A1", "L1 Ark1!B1", "L2 Ark1!C1"]);
    expect(calls.direct).toBe(3);
    expect(g.notFollowed).toEqual([]);

    // En fejl, der kører i ring, har ingen kilde.
    const ring = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ Ark1: { A1: ["=B1", "#REF!"], B1: ["=A1", "#REF!"] } }).reader);
    expect(cells(ring).some((c) => c.origin)).toBe(false);
    expect(renderTrace(ring)).toContain("The add-in could not determine in which cell the error starts.");
  });

  it("flere ark: arknavne med mellemrum og apostrof, og samme celle nået ad to veje står én gang", async () => {
    const { reader } = fakeReader({
      Oversigt: { A1: ["='Mit ark'!B2+'Bob''s'!C3+Data!A1", 6] },
      "Mit ark": { B2: ["=Data!A1*2", 2] },
      "Bob's": { C3: 3 },
      Data: { A1: 1 },
    });
    const g = await traceGraph({ sheet: "Oversigt", address: "A1" }, "precedents", reader);
    expect(names(g)).toEqual(["L0 Oversigt!A1", "L1 'Mit ark'!B2", "L1 'Bob''s'!C3", "L1 Data!A1"]);
    expect(renderTrace(g)).toContain("L1 'Mit ark'!B2: =Data!A1*2 → 2");
  });
});

describe("loft og beskæring (#76)", () => {
  /** En kæde: A1 = A2+1, A2 = A3+1, … den sidste er et tal. */
  const chain = (n: number): Book => ({
    Ark1: Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, i === n - 1 ? 1 : ([`=A${i + 2}+1`, n - i] as Cell)])),
  });

  it("dybdeloftet: gennemløbet stopper, siger hvor, og Excels samlede graf (ExcelApi 1.14) fortæller, hvad der mangler", async () => {
    const { reader, calls } = fakeReader(chain(30), "precedents", { whole: true });
    const limits = { ...TRACE_LIMITS, maxDepth: 5 };
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader, limits);
    expect(cells(g)).toHaveLength(6);
    expect(g.notFollowed).toEqual(["Ark1!A6"]);
    expect(calls.direct).toBe(5);
    expect(calls.whole).toBe(1);
    const text = renderTrace(g, limits);
    expect(text).toContain(
      "TRUNCATED: the trace stopped here (limit: 5 levels, 200 cells). Not traced further: Ark1!A6. The cells they are calculated from are not in the list.",
    );
    expect(text).toContain("According to Excel, the whole graph of precedents covers 29 cells in: Ark1!A2, Ark1!A3");
    expect(text).toContain("and 17 more ranges");
  });

  it("uden kaldet for hele grafen (før ExcelApi 1.14) siges det stadig, at grafen er beskåret", async () => {
    const { reader } = fakeReader(chain(30));
    const limits = { ...TRACE_LIMITS, maxDepth: 5 };
    const text = renderTrace(await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader, limits), limits);
    expect(text).toContain("TRUNCATED: the trace stopped here");
    expect(text).not.toContain("According to Excel");
  });

  it("en ubeskåret graf spørger ikke om hele grafen", async () => {
    const { reader, calls } = fakeReader(chain(5), "precedents", { whole: true });
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader);
    expect(cells(g)).toHaveLength(5);
    expect(calls.whole).toBe(0);
    expect(renderTrace(g)).not.toContain("TRUNCATED");
  });

  it("loftet over celler: højst 200 i listen, fejl beholdes før formler før værdier, og resten tælles", async () => {
    // 60 enkeltceller pr. formel i fire formler: 240 celler på niveau 2, heraf 3 med fejl sidst.
    const book: Book = { Ark1: { A1: ["=B1+B2+B3+B4", 0] } };
    for (let f = 1; f <= 4; f++) {
      const refs = Array.from({ length: 60 }, (_, i) => `${colName(3 + i)}${f}`);
      book.Ark1[`B${f}`] = [`=${refs.join("+")}`, 0];
      refs.forEach((r, i) => (book.Ark1[r] = f === 4 && i >= 57 ? (["=1/0", "#DIV/0!"] as Cell) : i));
    }
    const { reader } = fakeReader(book);
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader);
    expect(cells(g)).toHaveLength(200);
    expect(g.omitted).toBe(45);
    // De tre fejlceller stod sidst, men er med.
    expect(cells(g).filter((c) => c.error)).toHaveLength(3);
    expect(g.omittedErrors).toBe(0);
    // Rækkefølgen er stadig den fundne.
    expect(cells(g)[5].address).toBe("D1");
    const text = renderTrace(g);
    expect(text).toContain("TRUNCATED: the list stops at 200 cells. 45 more cells were found but are not listed.");
    expect(text.length).toBeLessThanOrEqual(TRACE_LIMITS.maxChars);
  });

  it("når loftet er nået, følges resten ikke, og det siges", async () => {
    const book: Book = { Ark1: { A1: ["=B1:B5", 0] } };
    for (let i = 1; i <= 5; i++) {
      book.Ark1[`B${i}`] = [`=C${i}`, 0];
      book.Ark1[`C${i}`] = i;
    }
    const limits = { ...TRACE_LIMITS, maxCells: 4 };
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader(book).reader, limits);
    expect(names(g)).toEqual(["L0 Ark1!A1", "L1 Ark1!B1", "L1 Ark1!B2", "L1 Ark1!B3"]);
    expect(g.omitted).toBe(2);
    expect(g.notFollowed).toEqual(["Ark1!B1", "Ark1!B2", "Ark1!B3"]);
    expect(renderTrace(g, limits)).toContain("TRUNCATED: the list stops at 4 cells. 2 more cells were found but are not listed.");
  });

  it("områdenoten tæller det, der faktisk står i listen, også når loftet over celler skærer i området", async () => {
    const book: Book = { Ark1: { A1: ["=SUM(B1:B20)", 210] } };
    for (let i = 1; i <= 20; i++) book.Ark1[`B${i}`] = i;
    const limits = { ...TRACE_LIMITS, maxCells: 6 };
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader(book).reader, limits);
    expect(cells(g)).toHaveLength(6);
    expect(g.items.find((x) => x.type === "area")).toMatchObject({ cells: 20, listed: 5, constants: 8 });
    expect(g.omitted).toBe(7);
    expect(renderTrace(g, limits)).toContain("L1 Ark1!B1:B20: a range of 20 cells; 5 listed below, not listed (none of them has an error): 8 values.");
  });

  it("en celle uden formel har intet at spore", async () => {
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ Ark1: { A1: 42 } }).reader);
    expect(cells(g)).toHaveLength(1);
    expect(renderTrace(g)).toContain("L0 Ark1!A1: 42\nThe cell holds a value, not a formula, so nothing is traced.");
  });

  it("et stort område: fejlcellerne og de første andre vises, resten tælles i en note", async () => {
    const book: Book = { Ark1: { A1: ["=SUM(Data!B1:B100)", "#N/A"] }, Data: {} };
    for (let i = 1; i <= 100; i++) book.Data[`B${i}`] = i === 90 ? (["=VLOOKUP(1,C1:C2,1,FALSE)", "#N/A"] as Cell) : i % 2 ? i : ([`=B${i - 1}*2`, i] as Cell);
    delete book.Data.B50;
    const { reader } = fakeReader(book);
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader);
    const note = g.items.find((x) => x.type === "area");
    expect(note).toMatchObject({
      sheet: "Data",
      range: "B1:B100",
      depth: 1,
      cells: 100,
      listed: AREA_SAMPLE + 1,
      formulas: 42,
      constants: 44,
      empty: 1,
      unread: 0,
    });
    // Fejlcellen i række 90 er med, selv om den ligger efter de første tolv.
    expect(
      cells(g)
        .filter((c) => c.error)
        .map(refText),
    ).toEqual(["Ark1!A1", "Data!B90"]);
    const text = renderTrace(g);
    expect(text).toContain(
      "L1 Data!B1:B100: a range of 100 cells; 13 listed below, not listed (none of them has an error): 42 formulas, 44 values, 1 empty cell.",
    );
    // Fejlen følges videre fra B90, og dens egne celler er uden fejl.
    expect(cells(g).find((c) => c.address === "B90")).toMatchObject({ origin: true });
  });

  it("en hel kolonne læses kun i sin brugte del, og kun de første celler af et meget stort område undersøges", async () => {
    const book: Book = { Ark1: { A1: ["=SUM(Data!B:B)", 0] }, Data: {} };
    for (let i = 1; i <= MAX_AREA_READ + 500; i++) book.Data[`B${i}`] = i;
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader(book).reader);
    expect(g.items.find((x) => x.type === "area")).toMatchObject({ range: "B:B", cells: MAX_AREA_READ + 500, listed: AREA_SAMPLE, unread: 500 });
    expect(renderTrace(g)).toContain("TRUNCATED: 500 cells of the range were not examined.");
    // Kilden kan ikke afgøres, når en del af området ikke er undersøgt.
    const err = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader({ ...book, Ark1: { A1: ["=SUM(Data!B:B)", "#REF!"] } }).reader);
    expect(cells(err).some((c) => c.origin)).toBe(false);
  });

  it("teksten holder sig under loftet over tegn og siger, hvor mange linjer der mangler", async () => {
    const book: Book = { Ark1: { A1: [`=${Array.from({ length: 60 }, (_, i) => `${colName(1 + i)}2`).join("+")}`, 0] } };
    for (let i = 0; i < 60; i++) book.Ark1[`${colName(1 + i)}2`] = "En længere tekst i cellen, som fylder en del i listen til modellen";
    const limits = { ...TRACE_LIMITS, maxChars: 2_000 };
    const text = renderTrace(await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", fakeReader(book).reader, limits), limits);
    expect(text.length).toBeLessThanOrEqual(2_000);
    expect(text).toMatch(/TRUNCATED: \d+ lines were left out to keep the list short\./);
  });

  it("det, Excel ikke kan svare på, nævnes – og giver ingen kilde", async () => {
    const { reader } = fakeReader({ Ark1: { A1: ["=B1+Slettet!A1", "#REF!"], B1: 1 } }, "precedents", { fail: [] });
    const g = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", reader);
    expect(g.failed).toEqual(["Slettet!A1"]);
    expect(cells(g)[0].origin).toBeUndefined();
    expect(renderTrace(g)).toContain("Excel could not trace or read: Slettet!A1.");

    const failing = fakeReader({ Ark1: { A1: ["=B1", "#REF!"], B1: 1 } }, "precedents", { fail: ["Ark1!A1"] });
    const h = await traceGraph({ sheet: "Ark1", address: "A1" }, "precedents", failing.reader);
    expect(h.failed).toEqual(["Ark1!A1"]);
    expect(cells(h)).toHaveLength(1);
    expect(cells(h)[0].origin).toBeUndefined();
  });

  it("en celle, der ikke kan læses, giver en fejl med adressen", async () => {
    await expect(traceGraph({ sheet: "Findes ikke", address: "A1" }, "precedents", fakeReader({ Ark1: {} }).reader)).rejects.toThrow(
      "the cell 'Findes ikke'!A1 could not be read",
    );
  });
});

describe("afhængige (#76)", () => {
  it("cellerne, der bruger en celle, niveau for niveau og på tværs af ark", async () => {
    const { reader, calls } = fakeReader(DIV, "dependents");
    const g = await traceGraph({ sheet: "Data", address: "B2" }, "dependents", reader);
    expect(names(g)).toEqual(["L0 Data!B2", "L1 Beregning!B4", "L2 Beregning!B5", "L3 Resultat!B2", "L4 Resultat!B3"]);
    // Afhængige følges uanset fejl, og ingen celle er "kilde".
    expect(g.followedErrors).toBe(false);
    expect(cells(g).some((c) => c.origin)).toBe(false);
    expect(calls.direct).toBe(5);
    const text = renderTrace(g);
    expect(text).toContain("Dependents of Data!B2: the cells whose formulas use it");
    expect(text).toContain("L2 Beregning!B5: =B3/B4 → #DIV/0! [error]");
  });

  it("en celle, ingen formel bruger", async () => {
    const g = await traceGraph({ sheet: "Resultat", address: "A2" }, "dependents", fakeReader(DIV, "dependents").reader);
    expect(cells(g)).toHaveLength(1);
    expect(renderTrace(g)).toContain("No formula uses this cell.");
  });
});
