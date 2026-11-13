import { fitGrid, parseTsv, rangeBox, type AgentAction, type ApplyMode, type DocContext, type FormulaResult } from "@miai/shared";
import { mdToPlain } from "../markdown";
import { runExcelScript } from "../script";
import { FORMAT_PROPS } from "./excelFormats";
import type { AgentOptions, AgentOutcome, AgentStepResult, ApplyOutcome, ContextOptions, HostAdapter, TextFormat } from "./types";

const MAX_ROWS = 200;
const MAX_COLS = 30;
/** Statistikken regnes over højst dette antal rækker i used range (afgrænsning nævnes i prompten). */
const MAX_STAT_ROWS = 5_000;
/** Uden statistik: overskrifter + så mange rækker af used range. */
const HEADER_ROWS = 6;
/** Agentens overblik over projektmappen dækker højst så mange ark. */
const MAX_SHEETS = 30;

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
      const excel = opts.structure ? await this.structure(ctx, ws, used, sheetList.items, sheetUsed) : undefined;
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
   * Agent-tilstand: projektmappens ark (navn, brugt område, overskrifter) og Excel-tabeller (navn,
   * område med arknavn, overskrifter), og det aktive arks talformater i første datarække. Så kan
   * agenten arbejde på tværs af ark (#21), tilføje rækker i tabeller og skrive tal i det rigtige format.
   */
  private async structure(
    ctx: Excel.RequestContext,
    ws: Excel.Worksheet,
    used: Excel.Range,
    sheets: Excel.Worksheet[],
    sheetUsed: Excel.Range[],
  ): Promise<NonNullable<DocContext["excel"]>> {
    const tbls = ctx.workbook.tables;
    tbls.load("items/name");
    // Overskrifter: første række i hvert arks brugte område (højst MAX_SHEETS ark og MAX_COLS kolonner).
    const listed = sheets.slice(0, MAX_SHEETS);
    for (const r of sheetUsed.slice(0, MAX_SHEETS)) if (!r.isNullObject) r.load("columnCount");
    await ctx.sync();
    const headerRows = listed.map((_, i) => {
      const r = sheetUsed[i];
      if (r.isNullObject) return undefined;
      const h = r.getRow(0).getResizedRange(0, Math.min(r.columnCount, MAX_COLS) - r.columnCount);
      h.load("values");
      return h;
    });
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
    const text = (v: unknown) => String(v ?? "").slice(0, 255);
    return {
      tables: ranges.map((x) => ({ name: x.name, address: x.r.address, headers: (x.h.values[0] ?? []).map(text) })),
      columnFormats,
      activeSheet: ws.name,
      sheets: listed.map((sh, i) => ({
        name: sh.name,
        used: sheetUsed[i].isNullObject ? undefined : splitAddress(sheetUsed[i].address).range,
        headers: (headerRows[i]?.values[0] ?? []).map(text),
      })),
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
   * Agent-tilstand (#6, ADR-0012): udfører de bekræftede handlinger på markeringens ark og bevarer
   * formateringen:
   *  - værdier og formler ændrer aldrig cellernes format; tomme celler lige under eller ved siden af
   *    data får naboens talformat, skrift og fyld (ExcelApi 1.9);
   *  - tabeller får udseende med Excels indbyggede tabeltypografier (følger temaet), aldrig fri formatering;
   *  - rækker slettes til sidst og nedefra, så adresserne i planen (arket før handlingerne) holder;
   *    rækker i en Excel-tabel slettes som tabelrækker, så tabellen krymper;
   *  - nye tabelrækker tilføjes med tabellen selv (tabellens format og beregnede kolonner).
   * Fortryd gendanner alt i omvendt rækkefølge: slettede rækker sættes ind igen med værdier og
   * formater, tilføjede tabelrækker fjernes, nye tabeller bliver almindelige områder igen, og formler,
   * talformater, formater og kolonnebredder får deres tidligere værdier.
   */
  async applyAgent(actions: AgentAction[], dc: DocContext, opts: AgentOptions = {}): Promise<AgentOutcome> {
    if (!dc.address) throw new Error("Ingen markering.");
    const sheetName = splitAddress(dc.address).sheet;
    const script = actions.find((a) => a.action === "runScript");
    if (script?.action === "runScript") {
      if (!opts.scripts) throw new Error("Kodetilstanden er ikke slået til (Indstillinger → Avanceret).");
      if (actions.length > 1) throw new Error("Et script skal stå alene i planen.");
      return runExcelScript(script, sheetName);
    }

    // Flere ark (#21): handlingerne grupperes pr. ark og udføres ark for ark med samme fremgangsmåde
    // og fortryd. Nye ark oprettes først, så planen kan skrive på dem; omdøbning sker til sidst, så
    // alle adresser i planen bruger navnene fra før planen.
    const tableSheet = new Map((dc.excel?.tables ?? []).map((t) => [t.name, t.address.includes("!") ? splitAddress(t.address).sheet : sheetName]));
    const sheetFor = (a: AgentAction): string => {
      const ref = "address" in a ? a.address : "target" in a ? a.target : "source" in a ? a.source : "table" in a ? a.table : undefined;
      if (!ref) return sheetName;
      if (ADDRESS.test(ref)) return ref.includes("!") ? splitAddress(ref).sheet : sheetName;
      return tableSheet.get(ref) ?? sheetName;
    };
    const results: AgentStepResult[] = actions.map(() => ({ ok: true }));
    const message = (e: unknown) => (e instanceof Error ? e.message : String(e));

    const added: string[] = [];
    if (actions.some((a) => a.action === "addSheet"))
      await Excel.run(async (ctx) => {
        for (const [i, a] of actions.entries()) {
          if (a.action !== "addSheet") continue;
          try {
            ctx.workbook.worksheets.add(a.name);
            await ctx.sync();
            added.push(a.name);
          } catch (e) {
            results[i] = { ok: false, error: "failed", detail: message(e) };
          }
        }
      });

    const structural = (a: AgentAction) => a.action === "addSheet" || a.action === "renameSheet";
    const order = [...new Set(actions.filter((a) => !structural(a)).map(sheetFor))];
    const undos: (() => Promise<void>)[] = [];
    for (const name of order) {
      const idx = actions.flatMap((a, i) => (!structural(a) && sheetFor(a) === name ? [i] : []));
      try {
        const out = await this.applyOnSheet(
          idx.map((i) => actions[i]),
          dc,
          opts,
          name,
          name === sheetName,
        );
        idx.forEach((i, k) => (results[i] = out.results[k]));
        // Et ark, planen selv har oprettet, fjernes helt ved fortryd.
        if (!added.includes(name)) undos.push(out.undo);
      } catch (e) {
        for (const i of idx) results[i] = { ok: false, error: "failed", detail: message(e) };
      }
    }

    const renamed: { from: string; to: string }[] = [];
    if (actions.some((a) => a.action === "renameSheet"))
      await Excel.run(async (ctx) => {
        for (const [i, a] of actions.entries()) {
          if (a.action !== "renameSheet") continue;
          try {
            ctx.workbook.worksheets.getItem(a.sheet).name = a.name;
            await ctx.sync();
            renamed.push({ from: a.sheet, to: a.name });
          } catch (e) {
            results[i] = { ok: false, error: "failed", detail: message(e) };
          }
        }
      });

    const undo = async () => {
      // Omvendt rækkefølge: navnene tilbage, så arkene kan findes; så hvert arks ændringer; så de nye ark væk.
      if (renamed.length)
        await Excel.run(async (ctx) => {
          for (const r of [...renamed].reverse()) ctx.workbook.worksheets.getItem(r.to).name = r.from;
          await ctx.sync();
        });
      for (const u of [...undos].reverse()) await u();
      if (added.length)
        await Excel.run(async (ctx) => {
          for (const name of [...added].reverse()) ctx.workbook.worksheets.getItem(name).delete();
          await ctx.sync();
        });
    };
    return { results, undo, undoKind: "cells" };
  }

  /** Udfører handlingerne for ét ark (se applyAgent). `active`: arket er det, markeringen står på. */
  private async applyOnSheet(actions: AgentAction[], dc: DocContext, opts: AgentOptions, sheetName: string, active: boolean): Promise<AgentOutcome> {
    const has = (v: string) => Office.context.requirements.isSetSupported("ExcelApi", v);
    const formats = has("1.9");
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const results: AgentStepResult[] = actions.map(() => ({ ok: true }));
      const fail = (i: number, detail: string) => (results[i] = { ok: false, error: "failed", detail });
      /** Adresse på arket (uden arknavn), eller undefined hvis den peger på et andet ark. */
      const local = (i: number, address: string): string | undefined => {
        const { sheet: s, range } = address.includes("!") ? splitAddress(address) : { sheet: sheetName, range: address };
        if (s !== sheetName) {
          results[i] = { ok: false, error: "other-sheet" };
          return undefined;
        }
        return range.replace(/\$/g, "");
      };
      // Kun tabellerne på dette ark (tabeladresser har arknavn; uden arknavn gælder det aktive ark).
      const knownTables = (dc.excel?.tables ?? [])
        .filter((t) => (t.address.includes("!") ? splitAddress(t.address).sheet === sheetName : active))
        .map((t) => ({ name: t.name, box: rangeBox(t.address) }));
      const isTableName = (target: string) => !ADDRESS.test(target);

      // 1) Find målområderne. formatTable på et område gemmer også områdets formater (til fortryd).
      const targets = actions.map((a, i): Excel.Range | undefined => {
        switch (a.action) {
          case "setValues": {
            const r = local(i, a.address);
            if (!r) return undefined;
            const cols = Math.max(...a.values.map((row) => row.length));
            return sheet
              .getRange(r)
              .getCell(0, 0)
              .getResizedRange(a.values.length - 1, cols - 1);
          }
          case "setFormula": {
            const r = local(i, a.address);
            return r ? sheet.getRange(r).getCell(0, 0) : undefined;
          }
          case "clear":
          case "setNumberFormat":
          case "autofitColumns": {
            const r = local(i, a.address);
            return r ? sheet.getRange(r) : undefined;
          }
          case "formatTable": {
            if (isTableName(a.target)) return undefined;
            if (!formats) {
              fail(i, "Kræver ExcelApi 1.9 (for at kunne fortryde formateringen)");
              return undefined;
            }
            const r = local(i, a.target);
            return r ? sheet.getRange(r) : undefined;
          }
          case "sortRange": {
            if (isTableName(a.target)) return sheet.tables.getItem(a.target).getRange();
            const r = local(i, a.target);
            return r ? sheet.getRange(r) : undefined;
          }
          case "conditionalFormat": {
            if (!has("1.6")) {
              fail(i, "Kræver ExcelApi 1.6");
              return undefined;
            }
            const r = local(i, a.address);
            return r ? sheet.getRange(r) : undefined;
          }
          case "freezeRows":
            if (!has("1.7")) fail(i, "Kræver ExcelApi 1.7");
            return undefined;
          case "saveSheet":
            if (!opts.exportSheet) fail(i, "Eksport kræver MIAI Bridge (Indstillinger → Eksport)");
            return undefined;
          case "appendRows":
          case "deleteRows":
          case "deleteEmptyRows":
          case "addChart":
            return undefined;
          default:
            fail(i, `${a.action} findes ikke i Excel`);
            return undefined;
        }
      });
      for (const t of targets) t?.load("address,formulas,numberFormat,rowIndex,columnIndex,rowCount,columnCount");
      // Eksisterende tabeller, der skal have en ny typografi: den gamle gemmes til fortryd.
      const styled = actions.map((a) => (a.action === "formatTable" && isTableName(a.target) ? sheet.tables.getItem(a.target) : undefined));
      for (const t of styled) t?.load("name,style");
      const appendTables = actions.map((a) => (a.action === "appendRows" ? sheet.tables.getItem(a.table) : undefined));
      for (const t of appendTables) t?.rows.load("count");
      // Frysning før ændringen (til fortryd).
      const freezeBefore = actions.some((a, i) => a.action === "freezeRows" && results[i].ok) ? sheet.freezePanes.getLocationOrNullObject() : undefined;
      freezeBefore?.load("address");
      // Kolonnebredder før autofit (højst 50 kolonner pr. handling).
      const widths = actions.map((a, i) => {
        if (a.action !== "autofitColumns" || !targets[i]) return undefined;
        const b = rangeBox(a.address);
        const n = Math.min(b.c1 - b.c0 + 1, 50);
        return Array.from({ length: n }, (_, c) => {
          const column = targets[i]!.getColumn(c);
          column.format.load("columnWidth");
          return column;
        });
      });
      await ctx.sync();
      // Naboer til tomme mål for nye værdier: rækken over og kolonnen til venstre (samme størrelse).
      const writes = (k: number) => actions[k].action === "setValues" || actions[k].action === "setFormula";
      const neighbours = targets.map((t, i) => {
        if (!t || !formats || !writes(i)) return undefined;
        const up = t.rowIndex > 0 ? t.getOffsetRange(-1, 0) : undefined;
        const left = t.columnIndex > 0 ? t.getOffsetRange(0, -1) : undefined;
        up?.load("formulas");
        left?.load("formulas");
        return { up, left };
      });
      const props = targets.map((t) => (t && formats ? t.getCellProperties(FORMAT_PROPS) : undefined));
      await ctx.sync();

      const snapshots = targets.flatMap((t) =>
        t
          ? [
              {
                address: t.address,
                formulas: t.formulas,
                numberFormat: t.numberFormat,
                props: props[targets.indexOf(t)]?.value as Excel.SettableCellProperties[][] | undefined,
              },
            ]
          : [],
      );
      const added = appendTables.flatMap((t, i) =>
        t ? [{ name: (actions[i] as Extract<AgentAction, { action: "appendRows" }>).table, from: t.rows.count }] : [],
      );
      const oldStyles = styled.flatMap((t) => (t ? [{ name: t.name, style: t.style }] : []));
      const oldWidths = widths.flatMap((cols, i) =>
        cols ? cols.map((c, k) => ({ address: targets[i]!.address, column: k, width: c.format.columnWidth })) : [],
      );
      const isEmpty = (f: unknown[][]) => f.every((row) => row.every((v) => v === "" || v === null));
      const hasData = (r?: Excel.Range) => !!r && !isEmpty(r.formulas);

      // 2) Værdier, formler, talformater, kolonnebredder, tabeltypografier, sortering, diagrammer …
      const created: { name?: string; table: Excel.Table; box: ReturnType<typeof rangeBox>; i: number }[] = [];
      const charts: { chart: Excel.Chart; name?: string }[] = [];
      const condFormats: { address: string; cf: Excel.ConditionalFormat; id?: string }[] = [];
      let froze = false;
      actions.forEach((a, i) => {
        const t = targets[i];
        if (!results[i].ok) return;
        if (a.action === "addChart") {
          const tbl = isTableName(a.source) ? knownTables.find((x) => x.name === a.source) : undefined;
          const r = tbl ? undefined : local(i, a.source);
          if (!tbl && !r) return;
          const box = tbl?.box ?? rangeBox(r!);
          const chart = sheet.charts.add(
            CHART_TYPES[a.type] as Excel.ChartType,
            tbl ? sheet.tables.getItem(a.source).getRange() : sheet.getRange(r!),
            Excel.ChartSeriesBy.auto,
          );
          if (a.title) chart.title.text = a.title;
          // Ved siden af data: én kolonne fri til højre, øverst ud for overskriften.
          chart.setPosition(sheet.getRange(`${colName(box.c1 + 2)}${box.r0 + 1}`));
          chart.load("name");
          charts.push({ chart });
          return;
        }
        if (a.action === "freezeRows") {
          if (a.rows === 0) sheet.freezePanes.unfreeze();
          else sheet.freezePanes.freezeRows(a.rows);
          froze = true;
          return;
        }
        if (a.action === "formatTable" && styled[i]) {
          styled[i]!.style = a.style;
          return;
        }
        if (a.action === "appendRows") {
          const width = (dc.excel?.tables.find((x) => x.name === a.table)?.headers.length ?? Math.max(...a.values.map((r) => r.length))) || 1;
          // null = Excel udfylder selv (beregnede kolonner) / lader cellen være.
          appendTables[i]!.rows.add(
            undefined,
            a.values.map((row) => [...row, ...Array(Math.max(0, width - row.length)).fill(null)]) as (string | number | boolean)[][],
          );
          results[i] = { ok: true, count: a.values.length };
          return;
        }
        if (!t) return;
        if (writes(i)) {
          // Tomme celler får naboens format, før værdien skrives (så fx talformatet gælder med det samme).
          const n = neighbours[i];
          if (n && isEmpty(t.formulas)) {
            const src = hasData(n.up) ? n.up : hasData(n.left) ? n.left : undefined;
            if (src) t.copyFrom(src, Excel.RangeCopyType.formats);
          }
        }
        switch (a.action) {
          case "setValues": {
            const cols = Math.max(...a.values.map((row) => row.length));
            t.values = a.values.map((row) => [...row, ...Array(cols - row.length).fill(null)]) as (string | number | boolean)[][];
            results[i] = { ok: true, count: a.values.length * cols };
            break;
          }
          case "setFormula": {
            const f = a.formula.trim();
            t.formulas = [[f.startsWith("=") ? f : `=${f}`]];
            results[i] = { ok: true, count: 1 };
            break;
          }
          case "clear":
            t.clear(Excel.ClearApplyTo.contents);
            break;
          case "setNumberFormat":
            t.numberFormat = Array.from({ length: t.rowCount }, () => Array(t.columnCount).fill(a.format) as string[]);
            results[i] = { ok: true, count: t.rowCount * t.columnCount };
            break;
          case "autofitColumns":
            t.format.autofitColumns();
            break;
          case "formatTable": {
            const table = sheet.tables.add(t, true);
            table.style = a.style;
            table.load("name");
            created.push({ table, box: rangeBox(t.address), i });
            break;
          }
          case "sortRange": {
            // Kolonnen efter overskrift (uden forskel på store/små bogstaver) eller bogstav.
            const headers = (t.formulas[0] ?? []).map((h) => String(h).trim().toLowerCase());
            let key = headers.indexOf(a.by.trim().toLowerCase());
            if (key < 0 && /^[A-Z]{1,3}$/i.test(a.by.trim())) key = colIndex(a.by.trim().toUpperCase()) - t.columnIndex;
            if (key < 0 || key >= t.columnCount) {
              fail(i, `Kolonnen "${a.by}" findes ikke`);
              break;
            }
            const fields = [{ key, ascending: a.order !== "desc" }];
            if (isTableName(a.target)) sheet.tables.getItem(a.target).sort.apply(fields);
            else t.sort.apply(fields, false, true);
            break;
          }
          case "conditionalFormat": {
            const cf = addConditionalFormat(t, a);
            cf.load("id");
            condFormats.push({ address: t.address, cf });
            break;
          }
        }
      });
      await ctx.sync();
      // saveSheet (tillæg ADR-0012): arket læses og eksporteres som .xlsx via bridgen.
      // Det sker efter ændringerne, så filen viser arket med de udførte handlinger.
      for (let i = 0; i < actions.length; i++) {
        const a = actions[i];
        if (a.action !== "saveSheet" || !results[i].ok) continue;
        try {
          if (!opts.exportSheet) throw new Error("Eksport kræver MIAI Bridge (Indstillinger → Eksport)");
          const ws = a.sheet ? ctx.workbook.worksheets.getItem(a.sheet) : sheet;
          const used = ws.getUsedRangeOrNullObject(true);
          if (!used) throw new Error("Arket har ingen data");
          used.load("values,formulas,numberFormat,columnCount");
          const widths = Array.from({ length: Math.min(used.columnCount, 256) }, (_, c) => {
            const column = used.getColumn(c);
            column.format.load("columnWidth");
            return column;
          });
          await ctx.sync();
          if (used.isNullObject) throw new Error("Arket har ingen data");
          const path = await opts.exportSheet({
            fileName: a.fileName,
            folder: "desktop",
            name: a.sheet ?? sheetName,
            address: splitAddress(used.address).range,
            values: used.values as (string | number | boolean | null)[][],
            formulas: used.formulas as string[][],
            numberFormats: used.numberFormat as string[][],
            // Skjulte kolonner (bredde 0) springes over, så de får standardbredden i den nye fil.
            columnWidths: widths.map((c) => c.format.columnWidth).filter((w) => Number.isFinite(w) && w >= 1),
          });
          results[i] = { ok: true, count: 1, detail: path };
        } catch (e) {
          fail(i, e instanceof Error ? e.message : String(e));
        }
      }
      for (const c of created) c.name = c.table.name;
      for (const c of charts) c.name = c.chart.name;
      for (const c of condFormats) c.id = c.cf.id;

      // 3) Formelfejl (#VALUE!, #NAME? …) meldes pr. handling – før sletninger flytter cellerne.
      const formulaCells = actions.map((a, i) => (a.action === "setFormula" ? targets[i] : undefined));
      for (const c of formulaCells) c?.load("values,valueTypes");
      await ctx.sync();
      formulaCells.forEach((c, i) => {
        if (c && c.valueTypes[0][0] === Excel.RangeValueType.error) results[i] = { ok: false, error: "formula-error", detail: String(c.values[0][0]) };
      });

      // 4) Rækker, der skal slettes. Nye tabeller fra trin 2 tæller med som tabeller.
      const tablesNow = [...knownTables, ...created.map((c) => ({ name: c.name!, box: c.box }))];
      type Unit = { i: number; top: number } & ({ kind: "table"; table: string; index: number } | { kind: "range"; address: string });
      const units: Unit[] = [];
      const tableRows = (i: number, table: { name: string; box: ReturnType<typeof rangeBox> }, rows: number[]) =>
        rows.forEach((r) => units.push({ i, top: r, kind: "table", table: table.name, index: r - table.box.r0 - 1 }));
      const scans: { i: number; table?: { name: string; box: ReturnType<typeof rangeBox> }; range: Excel.Range; box?: ReturnType<typeof rangeBox> }[] = [];
      // Tomme rækker i et område tælles kun inden for arkets data (ikke de tomme rækker under det).
      const used = actions.some((a) => a.action === "deleteEmptyRows" && !isTableName(a.target)) ? sheet.getUsedRangeOrNullObject(true) : undefined;
      used?.load("rowIndex,rowCount");
      actions.forEach((a, i) => {
        if (a.action === "deleteRows") {
          const r = local(i, a.address);
          if (!r) return;
          const b = rangeBox(r);
          const table = tablesNow.find((t) => b.r0 > t.box.r0 && b.r1 <= t.box.r1 && b.c0 <= t.box.c1 && t.box.c0 <= b.c1);
          if (table) tableRows(i, table, range(b.r0, b.r1));
          else units.push({ i, top: b.r0, kind: "range", address: r });
        } else if (a.action === "deleteEmptyRows") {
          const target = isTableName(a.target) ? a.target : local(i, a.target);
          if (!target) return;
          const b = isTableName(target) ? undefined : rangeBox(target);
          // Et område, der rammer en tabel, behandles som tabellen (tabellens tomme rækker).
          const table = tablesNow.find((t) => (b ? b.r0 <= t.box.r1 && t.box.r0 <= b.r1 && b.c0 <= t.box.c1 && t.box.c0 <= b.c1 : t.name === target));
          if (!table && !b) {
            fail(i, `Tabellen ${target} findes ikke`);
            return;
          }
          const scan = table ? sheet.tables.getItem(table.name).getDataBodyRange() : sheet.getRange(target);
          scan.load("formulas");
          scans.push({ i, table, range: scan, box: table ? { ...table.box, r0: table.box.r0 + 1 } : b });
        }
      });
      await ctx.sync();
      const lastUsed = used ? (used.isNullObject ? -1 : used.rowIndex + used.rowCount - 1) : Infinity;
      for (const s of scans) {
        const empty = s.range.formulas.flatMap((row, k) => (isEmpty([row]) ? [s.box!.r0 + k] : [])).filter((r) => r <= lastUsed);
        if (s.table) tableRows(s.i, s.table, empty);
        else
          for (const [from, to] of groups(empty)) {
            units.push({ i: s.i, top: from, kind: "range", address: `${colName(s.box!.c0)}${from + 1}:${colName(s.box!.c1)}${to + 1}` });
          }
      }
      // Nedefra og hver række kun én gang, så de øvrige adresser og tabelindeks holder.
      const seen = new Set<string>();
      const ordered = units
        .sort((x, y) => y.top - x.top)
        .filter((u) => {
          const k = u.kind === "table" ? `${u.table}:${u.index}` : u.address;
          return !seen.has(k) && !!seen.add(k);
        });
      const rowRange = (c: Excel.RequestContext, ws: Excel.Worksheet, u: Unit) =>
        u.kind === "table" ? ws.tables.getItem(u.table).rows.getItemAt(u.index).getRange() : ws.getRange(u.address);
      const deleted: (Unit & { formulas: unknown[][]; numberFormat: string[][]; props?: Excel.SettableCellProperties[][] })[] = [];
      for (const u of ordered) {
        const r = rowRange(ctx, sheet, u);
        r.load("formulas,numberFormat");
        const p = formats ? r.getCellProperties(FORMAT_PROPS) : undefined;
        try {
          await ctx.sync();
          // Gem værdierne, før rækken forsvinder.
          const snap = { formulas: r.formulas, numberFormat: r.numberFormat, props: p?.value as Excel.SettableCellProperties[][] | undefined };
          if (u.kind === "table") sheet.tables.getItem(u.table).rows.getItemAt(u.index).delete();
          else r.delete(Excel.DeleteShiftDirection.up);
          await ctx.sync();
          deleted.push({ ...u, ...snap });
        } catch (e) {
          fail(u.i, e instanceof Error ? e.message : String(e));
        }
      }
      actions.forEach((a, i) => {
        if ((a.action === "deleteRows" || a.action === "deleteEmptyRows") && results[i].ok) {
          const n = deleted.filter((d) => d.i === i).reduce((sum, d) => sum + (d.kind === "table" ? 1 : d.formulas.length), 0);
          results[i] = { ok: true, count: n };
        }
      });

      const undo = () =>
        Excel.run(async (c) => {
          const ws = c.workbook.worksheets.getItem(sheetName);
          // Slettede rækker sættes ind igen oppefra (omvendt af sletningen) med værdier og formater.
          for (const d of [...deleted].reverse()) {
            if (d.kind === "table") ws.tables.getItem(d.table).rows.add(d.index, d.formulas as (string | number | boolean)[][]);
            else ws.getRange(d.address).insert(Excel.InsertShiftDirection.down);
            await c.sync();
            const r = rowRange(c, ws, d);
            r.formulas = d.formulas as string[][];
            r.numberFormat = d.numberFormat;
            if (d.props) r.setCellProperties(d.props);
            await c.sync();
          }
          // Nye tabeller bliver almindelige områder igen; tabeltypografiens formater ryddes (cellerne gendannes nedenfor).
          for (const t of [...created].reverse()) {
            if (!t.name) continue;
            ws.tables.getItem(t.name).convertToRange();
            ws.getRange(`${colName(t.box.c0)}${t.box.r0 + 1}:${colName(t.box.c1)}${t.box.r1 + 1}`).clear(Excel.ClearApplyTo.formats);
          }
          for (const s of oldStyles) ws.tables.getItem(s.name).style = s.style;
          // Cellerne (på de adresser, de havde under sessionen).
          for (const snap of [...snapshots].reverse()) {
            const r = ws.getRange(splitAddress(snap.address).range);
            r.formulas = snap.formulas as string[][];
            r.numberFormat = snap.numberFormat;
            if (snap.props) r.setCellProperties(snap.props);
          }
          for (const w of oldWidths) ws.getRange(splitAddress(w.address).range).getColumn(w.column).format.columnWidth = w.width;
          // Diagrammer og betinget formatering, agenten har lavet, fjernes; frysningen sættes tilbage.
          for (const ch of charts) if (ch.name) ws.charts.getItem(ch.name).delete();
          for (const cf of condFormats) if (cf.id) ws.getRange(splitAddress(cf.address).range).conditionalFormats.getItem(cf.id).delete();
          if (froze) {
            if (!freezeBefore || freezeBefore.isNullObject) ws.freezePanes.unfreeze();
            else ws.freezePanes.freezeAt(ws.getRange(splitAddress(freezeBefore.address).range));
          }
          await c.sync();
          // Tilføjede tabelrækker (bagfra, så indeks ikke flytter sig).
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

/** En adresse som "B2" eller "Ark1!B2:D9" (ellers er det et tabelnavn). */
const ADDRESS = /^(?:(?:'[^']+'|[^!'\s]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i;

const range = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, k) => from + k);

/** Sorterede rækkenumre → sammenhængende stykker [fra, til]. */
function groups(rows: number[]): [number, number][] {
  const out: [number, number][] = [];
  for (const r of rows) {
    const last = out[out.length - 1];
    if (last && r === last[1] + 1) last[1] = r;
    else out.push([r, r]);
  }
  return out;
}

/** Agentens diagramtyper → Excels. */
const CHART_TYPES: Record<string, string> = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "Line",
  pie: "Pie",
  area: "Area",
  scatter: "XYScatter",
};

/**
 * Excels standardformater til betinget formatering: farveskala (rød-gul-grøn), datalinjer og
 * "lys rød udfyldning med mørkerød tekst" til fremhævning – de samme forvalg som i Excels menu.
 */
function addConditionalFormat(r: Excel.Range, a: Extract<AgentAction, { action: "conditionalFormat" }>): Excel.ConditionalFormat {
  const highlight = (f: { fill: { color: string }; font: { color: string } }) => {
    f.fill.color = "#FFC7CE";
    f.font.color = "#9C0006";
  };
  switch (a.rule) {
    case "colorScale": {
      const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      cf.colorScale.criteria = {
        minimum: { formula: undefined, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
        midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percentile, color: "#FFEB84" },
        maximum: { formula: undefined, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" },
      };
      return cf;
    }
    case "dataBar":
      return r.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    case "greaterThan":
    case "lessThan": {
      const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      highlight(cf.cellValue.format);
      cf.cellValue.rule = { formula1: String(a.value ?? 0), operator: a.rule === "greaterThan" ? "GreaterThan" : "LessThan" };
      return cf;
    }
    case "duplicates": {
      const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
      highlight(cf.preset.format);
      cf.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.duplicateValues };
      return cf;
    }
    case "top10": {
      const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
      highlight(cf.topBottom.format);
      cf.topBottom.rule = { rank: 10, type: "TopItems" };
      return cf;
    }
  }
}
