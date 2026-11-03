import { fitGrid, parseTsv, type ApplyMode, type DocContext, type FormulaResult } from "@miai/shared";
import { mdToPlain } from "../markdown";
import type { ApplyOutcome, ContextOptions, HostAdapter, TextFormat } from "./types";

const MAX_ROWS = 200;
const MAX_COLS = 30;
/** Statistikken regnes over højst dette antal rækker i used range (afgrænsning nævnes i prompten). */
const MAX_STAT_ROWS = 5_000;

const toTsv = (values: unknown[][]) => values.map((r) => r.map((v) => String(v ?? "").replace(/[\t\n\r]+/g, " ")).join("\t")).join("\n");

/** 0 → "A", 26 → "AA" */
const colName = (index: number): string => {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    name = String.fromCharCode(65 + r) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
};

/** "C7" (eller "C7:E9") → 2, 0-baseret kolonneindex */
const colIndex = (address: string): number => {
  const m = /^([A-Z]+)/.exec(address);
  if (!m) return 0;
  let n = 0;
  for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

const fmtNum = (x: number): string => x.toFixed(2).replace(/\.?0+$/, "") || "0";

interface ColumnStat {
  letter: string;
  nonEmpty: number;
  numeric?: { sum: number; min: number; max: number };
}

/** Tæller ikke-tomme celler og summer/min/max pr. kolonne (kun ægte tal; datoer tæller som tekst). */
function columnStats(values: unknown[][], baseCol: number): ColumnStat[] {
  const cols = values[0]?.length ?? 0;
  const stats: ColumnStat[] = Array.from({ length: cols }, (_, c) => ({ letter: colName(baseCol + c), nonEmpty: 0 }));
  for (const row of values) {
    for (let c = 0; c < cols && c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined || v === "") continue;
      stats[c].nonEmpty++;
      if (typeof v === "number" && Number.isFinite(v)) {
        const n = stats[c].numeric ?? (stats[c].numeric = { sum: 0, min: v, max: v });
        n.sum += v;
        if (v < n.min) n.min = v;
        if (v > n.max) n.max = v;
      }
    }
  }
  return stats;
}

/** Kompakt oversigt over filen: arkliste, headers og kolonne-statistik regnet over hele used range. */
function buildOverview(
  sheetLines: string[],
  activeName: string,
  usedAddress: string, // fx "A1:E30" – uden arknavn
  usedRows: number,
  usedCols: number,
  values: unknown[][], // første MAX_STAT_ROWS rækker × maks. MAX_COLS kolonner
): string {
  const lines: string[] = [];
  lines.push(`Active sheet '${activeName}', used range ${usedAddress} (${usedRows} rows x ${usedCols} columns)`);
  if (sheetLines.length > 1) lines.push(`All sheets: ${sheetLines.join(" | ")}`);

  const stats = columnStats(values, colIndex(usedAddress));
  const firstRow = values[0] ?? [];
  if (firstRow.some((v) => typeof v === "string" && v.trim() !== "")) {
    lines.push(`Headers (row 1): ${stats.map((s, i) => `${s.letter}=${String(firstRow[i] ?? "").trim() || "-"}`).join(" ")}`);
  }
  lines.push(
    `Column stats (${values.length} rows): ${stats.map((s) => `${s.letter}: ${s.nonEmpty} non-empty${s.numeric ? `, numeric sum=${fmtNum(s.numeric.sum)} min=${fmtNum(s.numeric.min)} max=${fmtNum(s.numeric.max)}` : ""}`).join(" ")}`,
  );
  if (values.length < usedRows) lines.push(`Stats cover the first ${values.length} of ${usedRows} rows.`);
  const statCols = values[0]?.length ?? 0;
  if (statCols < usedCols) lines.push(`Stats cover the first ${statCols} of ${usedCols} columns.`);

  return lines.join("\n");
}

/** "'Mit ark'!A1:B2" → { sheet: "Mit ark", range: "A1:B2" } */
export function splitAddress(address: string): { sheet: string; range: string } {
  const i = address.lastIndexOf("!");
  const sheet = address.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
  return { sheet, range: address.slice(i + 1) };
}

export class ExcelAdapter implements HostAdapter {
  readonly host = "excel" as const;

  supports(mode: ApplyMode): boolean {
    if (mode === "comment") return Office.context.requirements.isSetSupported("ExcelApi", "1.10");
    return mode === "cells" || mode === "formula" || mode === "replace";
  }

  async getContext(_opts: ContextOptions): Promise<DocContext> {
    return Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load("address,rowCount,columnCount");

      const ws = sel.worksheet;
      ws.load("name");
      // Arkliste, så modellen kender filens struktur (fx "i tabellen" kan pege på et bestemt ark).
      const sheetList = ctx.workbook.worksheets;
      sheetList.load("items/name");

      // Used range i det aktive ark: basis for oversigt, headers og kolonne-statistik.
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("address,rowCount,columnCount");

      await ctx.sync();

      // Markering som TSV (værdier) – bruges til celle-for-celle-opgaver og chat.
      const rows = Math.min(sel.rowCount, MAX_ROWS);
      const cols = Math.min(sel.columnCount, MAX_COLS);
      const view = sel.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      view.load("values");
      const topLeft = sel.getCell(0, 0);
      topLeft.load("formulas");

      // Arkoversigt: used range pr. ark + værdier i det aktive arks used range (cappet).
      const sheetUsed = sheetList.items.map((s) => s.getUsedRangeOrNullObject(true));
      for (const r of sheetUsed) r.load("address");
      let data: Excel.Range | undefined;
      if (!used.isNullObject) {
        const dRows = Math.min(used.rowCount, MAX_STAT_ROWS);
        const dCols = Math.min(used.columnCount, MAX_COLS);
        data = used.getCell(0, 0).getResizedRange(dRows - 1, dCols - 1);
        data.load("values");
      }

      await ctx.sync();

      let surroundings: string | undefined;
      if (!used.isNullObject && data) {
        const values = (data.values as unknown[][]) ?? [];
        if (values.length > 0) {
          const sheetLines = sheetList.items.map((s, i) => `${s.name} ${sheetUsed[i].isNullObject ? "(empty)" : sheetUsed[i].address}`);
          // 20k tegn er DocContext.surroundings' grænse – afkortning sker der ved validering.
          surroundings = buildOverview(sheetLines, ws.name, used.address, used.rowCount, used.columnCount, values);
        }
      }

      const f = topLeft.formulas[0][0];
      const table = toTsv(view.values);
      return {
        host: "excel",
        address: sel.address,
        text: table,
        table,
        formula: typeof f === "string" && f.startsWith("=") ? f : undefined,
        surroundings,
        isWholeDocument: false,
      };
    });
  }

  private async withRange<T>(address: string, fn: (ctx: Excel.RequestContext, r: Excel.Range) => Promise<T>): Promise<T> {
    return Excel.run(async (ctx) => {
      const { sheet, range } = splitAddress(address);
      return fn(ctx, ctx.workbook.worksheets.getItem(sheet).getRange(range));
    });
  }

  /** Excel kan ikke altid fortryde add-in-ændringer med Ctrl+Z, så vi gemmer de tidligere formler. */
  private snapshotUndo(address: string, formulas: unknown[][]): () => Promise<void> {
    return () =>
      this.withRange(address, async (ctx, r) => {
        r.formulas = formulas as string[][];
        await ctx.sync();
      });
  }

  async applyText(text: string, format: TextFormat, mode: ApplyMode, dc: DocContext): Promise<ApplyOutcome> {
    if (!dc.address) throw new Error("Ingen markering.");
    const plain = format === "markdown" ? mdToPlain(text) : text;
    return this.withRange(dc.address, async (ctx, sel) => {
      if (mode === "comment") {
        const cell = sel.getCell(0, 0);
        cell.load("address");
        await ctx.sync();
        ctx.workbook.comments.add(cell.address, plain);
        await ctx.sync();
        return {};
      }
      sel.load("rowCount,columnCount");
      await ctx.sync();
      const single = mode === "replace" || sel.rowCount * sel.columnCount === 1 || format !== "tsv";
      const target = single ? sel.getCell(0, 0) : sel.getCell(0, 0).getResizedRange(Math.min(sel.rowCount, MAX_ROWS) - 1, Math.min(sel.columnCount, MAX_COLS) - 1);
      target.load("formulas,rowCount,columnCount,address");
      await ctx.sync();
      const undo = this.snapshotUndo(target.address, target.formulas);
      target.values = single ? [[plain]] : fitGrid(parseTsv(plain), target.rowCount, target.columnCount);
      await ctx.sync();
      return { undo };
    });
  }

  async applyFormula(result: FormulaResult, dc: DocContext): Promise<ApplyOutcome> {
    if (!dc.address) throw new Error("Ingen markering.");
    const formula = result.formula.trim().startsWith("=") ? result.formula.trim() : `=${result.formula.trim()}`;
    return this.withRange(dc.address, async (ctx, sel) => {
      const cell = sel.getCell(0, 0);
      cell.load("formulas,address");
      await ctx.sync();
      const undo = this.snapshotUndo(cell.address, cell.formulas);
      cell.formulas = [[formula]];
      await ctx.sync();
      // Validér: læs resultatet og se om Excel melder fejl.
      cell.load("values,valueTypes");
      await ctx.sync();
      if (cell.valueTypes[0][0] === Excel.RangeValueType.error) {
        return { undo, warning: true, message: `Formlen gav fejlen ${cell.values[0][0]}. Tjek den, eller fortryd.` };
      }
      return { undo, message: `Resultat: ${cell.values[0][0]}` };
    });
  }
}
