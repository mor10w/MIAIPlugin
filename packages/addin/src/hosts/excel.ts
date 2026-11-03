import { fitGrid, parseTsv, type AgentAction, type ApplyMode, type DocContext, type FormulaResult } from "@miai/shared";
import { mdToPlain } from "../markdown";
import type { AgentOutcome, AgentStepResult, ApplyOutcome, ContextOptions, HostAdapter, TextFormat } from "./types";

const MAX_ROWS = 200;
const MAX_COLS = 30;
/** Statistikken regnes over højst dette antal rækker i used range (afgrænsning nævnes i prompten). */
const MAX_STAT_ROWS = 5_000;
/** Uden statistik: overskrifter + så mange rækker af used range. */
const HEADER_ROWS = 6;

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

/**
 * Kompakt oversigt over filen: arkliste, headers og – med `withStats` – kolonne-statistik
 * regnet over hele used range. Uden statistik vises de første rækker i stedet.
 */
export function buildOverview(
  sheetLines: string[],
  activeName: string,
  usedAddress: string, // fx "A1:E30" – uden arknavn
  usedRows: number,
  usedCols: number,
  values: unknown[][], // første rækker × maks. MAX_COLS kolonner
  withStats = true,
): string {
  const lines: string[] = [];
  lines.push(`Active sheet '${activeName}', used range ${usedAddress} (${usedRows} rows x ${usedCols} columns)`);
  if (sheetLines.length > 1) lines.push(`All sheets: ${sheetLines.join(" | ")}`);

  const stats = columnStats(values, colIndex(usedAddress));
  const firstRow = values[0] ?? [];
  if (firstRow.some((v) => typeof v === "string" && v.trim() !== "")) {
    lines.push(`Headers (row 1): ${stats.map((s, i) => `${s.letter}=${String(firstRow[i] ?? "").trim() || "-"}`).join(" ")}`);
  }
  if (!withStats) {
    lines.push(`First rows:\n${toTsv(values)}`);
    return lines.join("\n");
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
  const sheet = address
    .slice(0, i)
    .replace(/^'(.*)'$/, "$1")
    .replace(/''/g, "'");
  return { sheet, range: address.slice(i + 1) };
}

export class ExcelAdapter implements HostAdapter {
  readonly host = "excel" as const;

  supports(mode: ApplyMode): boolean {
    if (mode === "comment") return Office.context.requirements.isSetSupported("ExcelApi", "1.10");
    return mode === "cells" || mode === "formula" || mode === "replace" || mode === "agent";
  }

  async getContext(opts: ContextOptions): Promise<DocContext> {
    const withStats = opts.overview === "stats";
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
        // Statistik kræver op til 5.000 × 30 celler; ellers hentes kun de første rækker.
        const dRows = Math.min(used.rowCount, withStats ? MAX_STAT_ROWS : HEADER_ROWS);
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
          // used.address indeholder arknavnet ("Ark1!A1:E30"); kolonnebogstaver skal regnes ud fra selve området.
          surroundings = buildOverview(sheetLines, ws.name, splitAddress(used.address).range, used.rowCount, used.columnCount, values, withStats);
        }
      }

      const f = topLeft.formulas[0][0];
      const table = toTsv(view.values);
      const excel = opts.structure ? await this.structure(ctx, ws, used) : undefined;
      return {
        ...(excel ? { excel } : {}),
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

  /**
   * Agent-tilstand: arkets Excel-tabeller (navn, område, overskrifter) og hver kolonnes talformat
   * i første datarække, så agenten kan tilføje rækker i tabellen og skrive tal i det rigtige format.
   */
  private async structure(ctx: Excel.RequestContext, ws: Excel.Worksheet, used: Excel.Range): Promise<NonNullable<DocContext["excel"]>> {
    const tbls = ws.tables;
    tbls.load("items/name");
    await ctx.sync();
    const ranges = tbls.items.map((t) => {
      const r = t.getRange();
      r.load("address");
      const h = t.getHeaderRowRange();
      h.load("values");
      return { name: t.name, r, h };
    });
    let firstData: Excel.Range | undefined;
    if (!used.isNullObject && used.rowCount > 1) {
      firstData = used.getRow(1).getResizedRange(0, Math.min(used.columnCount, MAX_COLS) - used.columnCount);
      firstData.load("numberFormat,columnIndex");
    }
    await ctx.sync();
    const columnFormats = firstData
      ? (firstData.numberFormat[0] ?? []).flatMap((fmt, i) => (fmt && fmt !== "General" ? [`${colName(firstData!.columnIndex + i)}: ${fmt}`] : []))
      : [];
    return {
      tables: ranges.map((x) => ({ name: x.name, address: x.r.address, headers: (x.h.values[0] ?? []).map((v) => String(v ?? "")) })),
      columnFormats,
    };
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
      const target = single
        ? sel.getCell(0, 0)
        : sel.getCell(0, 0).getResizedRange(Math.min(sel.rowCount, MAX_ROWS) - 1, Math.min(sel.columnCount, MAX_COLS) - 1);
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
  /**
   * Agent-tilstand (#6): skriver de bekræftede handlinger på markeringens ark. Alle berørte
   * områders formler gemmes først, så hele sessionen kan fortrydes (Excel kan ikke altid
   * fortryde add-in-ændringer med Ctrl+Z).
   */
  /**
   * Agent-tilstand (#6): skriver de bekræftede handlinger på markeringens ark og bevarer
   * formateringen:
   *  - værdier og formler ændrer aldrig cellernes format;
   *  - celler, der var tomme, får naboens format (rækken over eller kolonnen til venstre), så en
   *    ny total-række eller kolonne får samme talformat, skrift og fyld som data (ExcelApi 1.9);
   *  - rækker i en Excel-tabel tilføjes med tabellen selv, så de får tabellens formatering og
   *    beregnede kolonner.
   * Fortryd gendanner formler og (ExcelApi 1.9) formater i alle berørte områder og fjerner
   * tilføjede tabelrækker.
   */
  async applyAgent(actions: AgentAction[], dc: DocContext): Promise<AgentOutcome> {
    if (!dc.address) throw new Error("Ingen markering.");
    const sheetName = splitAddress(dc.address).sheet;
    const formats = Office.context.requirements.isSetSupported("ExcelApi", "1.9");
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const results: AgentStepResult[] = actions.map(() => ({ ok: true }));

      // 1) Find målområderne.
      const targets = actions.map((a, i): Excel.Range | undefined => {
        if (a.action === "appendRows") return undefined;
        if (a.action !== "setValues" && a.action !== "setFormula" && a.action !== "clear") {
          results[i] = { ok: false, error: "failed", detail: `${a.action} findes ikke i Excel` };
          return undefined;
        }
        const { sheet: s, range } = a.address.includes("!") ? splitAddress(a.address) : { sheet: sheetName, range: a.address };
        if (s !== sheetName) {
          results[i] = { ok: false, error: "other-sheet" };
          return undefined;
        }
        const r = sheet.getRange(range.replace(/\$/g, ""));
        if (a.action === "setValues") {
          const cols = Math.max(...a.values.map((row) => row.length));
          return r.getCell(0, 0).getResizedRange(a.values.length - 1, cols - 1);
        }
        return a.action === "setFormula" ? r.getCell(0, 0) : r;
      });
      const tables = actions.map((a) => (a.action === "appendRows" ? sheet.tables.getItem(a.table) : undefined));
      for (const t of targets) t?.load("address,formulas,numberFormat,rowIndex,columnIndex");
      for (const t of tables) t?.rows.load("count");
      // Naboer til tomme mål: rækken over og kolonnen til venstre (samme størrelse).
      const above = targets.map((t, i) => (t && actions[i].action !== "clear" ? t : undefined));
      await ctx.sync();
      const neighbours = above.map((t) => {
        if (!t || !formats) return undefined;
        const up = t.rowIndex > 0 ? t.getOffsetRange(-1, 0) : undefined;
        const left = t.columnIndex > 0 ? t.getOffsetRange(0, -1) : undefined;
        up?.load("formulas");
        left?.load("formulas");
        return { up, left };
      });
      // Formater gemmes til fortryd (ExcelApi 1.9).
      const FORMAT_PROPS = {
        format: {
          fill: { color: true },
          font: { bold: true, italic: true, color: true, name: true, size: true, underline: true, strikethrough: true },
          horizontalAlignment: true,
          verticalAlignment: true,
          wrapText: true,
          indentLevel: true,
        },
      } as Excel.CellPropertiesLoadOptions;
      const props = targets.map((t) => (t && formats ? t.getCellProperties(FORMAT_PROPS) : undefined));
      await ctx.sync();

      const snapshots = targets.flatMap((t, i) =>
        t
          ? [{ address: t.address, formulas: t.formulas, numberFormat: t.numberFormat, props: props[i]?.value as Excel.SettableCellProperties[][] | undefined }]
          : [],
      );
      const added = tables.flatMap((t, i) => (t ? [{ name: (actions[i] as Extract<AgentAction, { action: "appendRows" }>).table, from: t.rows.count }] : []));
      const isEmpty = (f: unknown[][]) => f.every((row) => row.every((v) => v === "" || v === null));
      const hasData = (r?: Excel.Range) => !!r && !isEmpty(r.formulas);

      // 2) Udfør.
      actions.forEach((a, i) => {
        const t = targets[i];
        if (a.action === "appendRows") {
          const tbl = tables[i]!;
          const width = (dc.excel?.tables.find((x) => x.name === a.table)?.headers.length ?? Math.max(...a.values.map((r) => r.length))) || 1;
          // null = Excel udfylder selv (beregnede kolonner) / lader cellen være.
          tbl.rows.add(undefined, a.values.map((row) => [...row, ...Array(Math.max(0, width - row.length)).fill(null)]) as (string | number | boolean)[][]);
          results[i] = { ok: true, count: a.values.length };
          return;
        }
        if (!t) return;
        // Tomme celler får naboens format, før værdien skrives (så fx talformatet gælder med det samme).
        const n = neighbours[i];
        if (n && isEmpty(t.formulas)) {
          const src = hasData(n.up) ? n.up : hasData(n.left) ? n.left : undefined;
          if (src) t.copyFrom(src, Excel.RangeCopyType.formats);
        }
        if (a.action === "setValues") {
          const cols = Math.max(...a.values.map((row) => row.length));
          t.values = a.values.map((row) => [...row, ...Array(cols - row.length).fill(null)]) as (string | number | boolean)[][];
          results[i] = { ok: true, count: a.values.length * cols };
        } else if (a.action === "setFormula") {
          const f = a.formula.trim();
          t.formulas = [[f.startsWith("=") ? f : `=${f}`]];
          results[i] = { ok: true, count: 1 };
        } else if (a.action === "clear") {
          t.clear(Excel.ClearApplyTo.contents);
          results[i] = { ok: true };
        }
      });
      await ctx.sync();

      // 3) Formelfejl (#VALUE!, #NAME? …) meldes pr. handling.
      const formulaCells = actions.map((a, i) => (a.action === "setFormula" ? targets[i] : undefined));
      for (const c of formulaCells) c?.load("values,valueTypes");
      await ctx.sync();
      formulaCells.forEach((c, i) => {
        if (c && c.valueTypes[0][0] === Excel.RangeValueType.error) results[i] = { ok: false, error: "formula-error", detail: String(c.values[0][0]) };
      });

      const undo = () =>
        this.withRange(dc.address!, async (c) => {
          const ws = c.workbook.worksheets.getItem(sheetName);
          // Først cellerne (på de adresser, de havde under sessionen) …
          for (const snap of [...snapshots].reverse()) {
            const r = ws.getRange(splitAddress(snap.address).range);
            r.formulas = snap.formulas as string[][];
            r.numberFormat = snap.numberFormat;
            if (snap.props) r.setCellProperties(snap.props);
          }
          await c.sync();
          // … så de tilføjede tabelrækker (bagfra, så indeks ikke flytter sig). Celler under tabellen rykker op igen.
          for (const t of [...added].reverse()) {
            const rows = ws.tables.getItem(t.name).rows;
            rows.load("count");
            await c.sync();
            for (let k = rows.count - 1; k >= t.from; k--) rows.getItemAt(k).delete();
          }
          await c.sync();
        });
      return { results, undo, undoKind: "cells" };
    });
  }
}
