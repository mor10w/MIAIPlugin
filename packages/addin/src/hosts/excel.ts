import { fitGrid, parseTsv, type ApplyMode, type DocContext, type FormulaResult } from "@miai/shared";
import { mdToPlain } from "../markdown";
import type { ApplyOutcome, ContextOptions, HostAdapter, TextFormat } from "./types";

const MAX_ROWS = 200;
const MAX_COLS = 30;

const toTsv = (values: unknown[][]) => values.map((r) => r.map((v) => String(v ?? "").replace(/[\t\n\r]+/g, " ")).join("\t")).join("\n");

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
      await ctx.sync();
      const rows = Math.min(sel.rowCount, MAX_ROWS);
      const cols = Math.min(sel.columnCount, MAX_COLS);
      const view = sel.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      view.load("values,formulas");
      const topLeft = sel.getCell(0, 0);
      topLeft.load("formulas");

      // Omgivelser: overskrifter og de første rækker af det brugte område, så modellen kender strukturen.
      const used = sel.worksheet.getUsedRangeOrNullObject(true);
      used.load("address,rowCount,columnCount");
      await ctx.sync();
      let surroundings: string | undefined;
      if (!used.isNullObject) {
        const head = used.getCell(0, 0).getResizedRange(Math.min(used.rowCount, 6) - 1, Math.min(used.columnCount, MAX_COLS) - 1);
        head.load("values,address");
        await ctx.sync();
        surroundings = `Used range: ${used.address}\nFirst rows (${head.address}):\n${toTsv(head.values)}`.slice(0, 4_000);
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
