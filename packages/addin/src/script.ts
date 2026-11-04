import { checkScript, isBannedKey, MAX_SCRIPT_ITERATIONS, type AgentAction } from "@miai/shared";
import { FORMAT_PROPS } from "./hosts/excelFormats";
import type { AgentOutcome } from "./hosts/types";

/**
 * Kodetilstand (ADR-0017): kører et Office.js-script fra agenten i Excel.
 *
 *  1. Scriptet tjekkes og omskrives af checkScript (kun `context`, `Excel` og sikre globale objekter).
 *  2. Det aktive ark gemmes som øjebliksbillede (celler, talformater og formater) sammen med listen
 *     over ark, tabeller, diagrammer og pivottabeller.
 *  3. Scriptet kører i strict mode med en løkkevagt og en tidsgrænse.
 *  4. Fejler det, rulles ændringerne straks tilbage, og fejlen kan sendes til modellen.
 *
 * Fortryd bruger samme øjebliksbillede. Det dækker ikke alt (fx ændringer på andre ark, slettede ark,
 * betinget formatering og datavalidering) – det står i UI'et og i ADR'en.
 */

/** Ark med flere celler end dette kan ikke gemmes som øjebliksbillede, og så køres scriptet ikke. */
export const MAX_SNAPSHOT_CELLS = 50_000;
/** Tidsgrænse for ét script. */
export const SCRIPT_TIMEOUT_MS = 60_000;

type ScriptFn = (context: Excel.RequestContext, excel: typeof Excel) => Promise<unknown>;

/** Oversætter det kontrollerede script til en funktion. Kaster, hvis sandkassen afviser det. */
export function compileScript(source: string): ScriptFn {
  const check = checkScript(source);
  if (!check.ok) throw new Error(check.error);
  // Koden er tjekket og omskrevet ovenfor; den kan kun nå parametrene og de sikre globale objekter.
  const fn = new Function("context", "Excel", "__key", "__tick", `"use strict";\nreturn (async () => {\n${check.code}\n})();`) as (
    c: Excel.RequestContext,
    e: typeof Excel,
    key: (k: unknown) => string | number,
    tick: () => void,
  ) => Promise<unknown>;
  return (context, excel) => {
    let n = 0;
    // Beregnede navne tjekkes ved kørsel og gøres til primitive værdier (en toString kan ikke snyde to gange).
    const key = (k: unknown) => {
      const p = typeof k === "number" ? k : String(k);
      if (isBannedKey(p)) throw new Error(`Egenskaben "${p}" er spærret i sandkassen.`);
      return p;
    };
    const tick = () => {
      if (++n > MAX_SCRIPT_ITERATIONS) throw new Error("Scriptet blev stoppet: for mange gennemløb (uendelig løkke?).");
    };
    return fn(context, excel, key, tick);
  };
}

interface SheetSnapshot {
  sheetName: string;
  sheets: string[];
  tables: string[];
  charts: string[];
  pivots: string[];
  used?: { address: string; formulas: unknown[][]; numberFormat: string[][]; props?: Excel.SettableCellProperties[][] };
}

const has = (v: string) => Office.context.requirements.isSetSupported("ExcelApi", v);

async function snapshotSheet(sheetName: string): Promise<SheetSnapshot> {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const ws = wb.worksheets.getItem(sheetName);
    wb.worksheets.load("items/name");
    wb.tables.load("items/name");
    ws.charts.load("items/name");
    const pivots = has("1.8") ? wb.pivotTables : undefined;
    pivots?.load("items/name");
    const used = ws.getUsedRangeOrNullObject();
    used.load("address,rowCount,columnCount,formulas,numberFormat");
    await ctx.sync();
    if (!used.isNullObject && used.rowCount * used.columnCount > MAX_SNAPSHOT_CELLS) {
      throw new Error(`Arket har over ${MAX_SNAPSHOT_CELLS.toLocaleString("da-DK")} celler. Kodetilstanden kører kun på mindre ark, så den kan fortrydes.`);
    }
    const props = !used.isNullObject && has("1.9") ? used.getCellProperties(FORMAT_PROPS) : undefined;
    await ctx.sync();
    return {
      sheetName,
      sheets: wb.worksheets.items.map((x) => x.name),
      tables: wb.tables.items.map((x) => x.name),
      charts: ws.charts.items.map((x) => x.name),
      pivots: pivots?.items.map((x) => x.name) ?? [],
      used: used.isNullObject
        ? undefined
        : {
            address: used.address,
            formulas: used.formulas,
            numberFormat: used.numberFormat,
            props: props?.value as Excel.SettableCellProperties[][] | undefined,
          },
    };
  });
}

/** Fjerner det, scriptet har tilføjet (ark, pivottabeller, diagrammer, tabeller), og gendanner arkets celler. */
async function restoreSheet(snap: SheetSnapshot): Promise<void> {
  await Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const ws = wb.worksheets.getItem(snap.sheetName);
    wb.worksheets.load("items/name");
    wb.tables.load("items/name");
    ws.charts.load("items/name");
    const pivots = has("1.8") ? wb.pivotTables : undefined;
    pivots?.load("items/name");
    await ctx.sync();
    for (const p of pivots?.items ?? []) if (!snap.pivots.includes(p.name)) p.delete();
    for (const c of ws.charts.items) if (!snap.charts.includes(c.name)) c.delete();
    for (const t of wb.tables.items) if (!snap.tables.includes(t.name)) t.convertToRange();
    await ctx.sync();
    for (const s of wb.worksheets.items) if (!snap.sheets.includes(s.name) && s.name !== snap.sheetName) s.delete();
    const now = ws.getUsedRangeOrNullObject();
    now.load("address");
    await ctx.sync();
    if (!now.isNullObject) {
      now.clear(Excel.ClearApplyTo.contents);
      now.clear(Excel.ClearApplyTo.formats);
    }
    if (snap.used) {
      const r = ws.getRange(snap.used.address.replace(/^.*!/, ""));
      r.formulas = snap.used.formulas as string[][];
      r.numberFormat = snap.used.numberFormat;
      if (snap.used.props) r.setCellProperties(snap.used.props);
    }
    await ctx.sync();
  });
}

/** Office-fejl med placering (fx "Range.load"), så modellen kan rette scriptet. */
function describeError(e: unknown): string {
  const err = e as { message?: string; code?: string; debugInfo?: { errorLocation?: string } };
  const where = err.debugInfo?.errorLocation ? ` (${err.debugInfo.errorLocation})` : "";
  return `${err.code && err.code !== err.message ? `${err.code}: ` : ""}${err.message ?? String(e)}${where}`;
}

export async function runExcelScript(a: Extract<AgentAction, { action: "runScript" }>, sheetName: string): Promise<AgentOutcome> {
  let fn: ScriptFn;
  try {
    fn = compileScript(a.code);
  } catch (e) {
    return {
      results: [{ ok: false, error: "failed", detail: (e as Error).message }],
      undo: async () => {},
      undoKind: "script",
      retry: { error: (e as Error).message, code: a.code },
    };
  }
  const snap = await snapshotSheet(sheetName);
  try {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error(`Scriptet blev stoppet efter ${SCRIPT_TIMEOUT_MS / 1000} sekunder.`)), SCRIPT_TIMEOUT_MS);
    });
    const value = await Promise.race([
      Excel.run(async (ctx) => {
        const r = await fn(ctx, Excel);
        await ctx.sync();
        return r;
      }),
      timeout,
    ]).finally(() => clearTimeout(timer));
    const detail = typeof value === "string" && value.trim() ? value.trim().slice(0, 300) : undefined;
    return { results: [{ ok: true, detail }], undo: () => restoreSheet(snap), undoKind: "script" };
  } catch (e) {
    const error = describeError(e);
    // Et halvt udført script efterlader ikke arket i en ukendt tilstand.
    let rolledBack = true;
    await restoreSheet(snap).catch(() => (rolledBack = false));
    return {
      results: [{ ok: false, error: "failed", detail: `${error}${rolledBack ? " – ændringerne er rullet tilbage" : ""}` }],
      undo: async () => {},
      undoKind: "script",
      retry: { error, code: a.code },
    };
  }
}
