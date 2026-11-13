import { checkScript, isBannedKey, MAX_SCRIPT_ITERATIONS, type AgentAction, type ScriptHost } from "@miai/shared";
import { FORMAT_PROPS } from "./hosts/excelFormats";
import type { AgentOutcome } from "./hosts/types";

/**
 * Kodetilstand (ADR-0017): kører et Office.js-script fra agenten i Excel, Word eller PowerPoint (#23).
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

type ScriptFn = (context: unknown, host: unknown) => Promise<unknown>;

/**
 * Oversætter det kontrollerede script til en funktion med `context` og værtens objekt (`Excel`,
 * `Word` eller `PowerPoint`) som eneste adgang til omverdenen. Kaster, hvis sandkassen afviser det.
 */
export function compileScript(source: string, host: ScriptHost = "Excel"): ScriptFn {
  const check = checkScript(source, host);
  if (!check.ok) throw new Error(check.error);
  // Koden er tjekket og omskrevet ovenfor; den kan kun nå parametrene og de sikre globale objekter.
  const fn = new Function("context", host, "__key", "__tick", `"use strict";\nreturn (async () => {\n${check.code}\n})();`) as (
    c: unknown,
    h: unknown,
    key: (k: unknown) => string | number,
    tick: () => void,
  ) => Promise<unknown>;
  return (context, hostObject) => {
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
    return fn(context, hostObject, key, tick);
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

/**
 * Fælles kørsel for alle værter: tjek og oversæt scriptet, tag et øjebliksbillede, kør med
 * tidsgrænse, og rul tilbage, hvis det fejler. `run` kører funktionen i værtens egen kontekst.
 */
async function runGuarded(
  a: Extract<AgentAction, { action: "runScript" }>,
  host: ScriptHost,
  snapshot: () => Promise<() => Promise<void>>,
  run: (fn: ScriptFn) => Promise<unknown>,
): Promise<AgentOutcome> {
  let fn: ScriptFn;
  try {
    fn = compileScript(a.code, host);
  } catch (e) {
    return {
      results: [{ ok: false, error: "failed", detail: (e as Error).message }],
      undo: async () => {},
      undoKind: "script",
      retry: { error: (e as Error).message, code: a.code },
    };
  }
  // Uden øjebliksbillede køres scriptet ikke (kaster, fx hvis arket er for stort eller API'et mangler).
  const restore = await snapshot();
  try {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error(`Scriptet blev stoppet efter ${SCRIPT_TIMEOUT_MS / 1000} sekunder.`)), SCRIPT_TIMEOUT_MS);
    });
    const value = await Promise.race([run(fn), timeout]).finally(() => clearTimeout(timer));
    const detail = typeof value === "string" && value.trim() ? value.trim().slice(0, 300) : undefined;
    return { results: [{ ok: true, detail }], undo: restore, undoKind: "script" };
  } catch (e) {
    const error = describeError(e);
    // Et halvt udført script efterlader ikke dokumentet i en ukendt tilstand.
    let rolledBack = true;
    await restore().catch(() => (rolledBack = false));
    return {
      results: [{ ok: false, error: "failed", detail: `${error}${rolledBack ? " – ændringerne er rullet tilbage" : ""}` }],
      undo: async () => {},
      undoKind: "script",
      retry: { error, code: a.code },
    };
  }
}

type RunScript = Extract<AgentAction, { action: "runScript" }>;

export function runExcelScript(a: RunScript, sheetName: string): Promise<AgentOutcome> {
  return runGuarded(
    a,
    "Excel",
    async () => {
      const snap = await snapshotSheet(sheetName);
      return () => restoreSheet(snap);
    },
    (fn) =>
      Excel.run(async (ctx) => {
        const r = await fn(ctx, Excel);
        await ctx.sync();
        return r;
      }),
  );
}

/**
 * Word (#23): øjebliksbilledet er brødtekstens OOXML, som sættes tilbage ved fejl og fortryd.
 * Sidehoved, sidefod og kommentarer er ikke med i brødteksten og gendannes ikke.
 */
export function runWordScript(a: RunScript): Promise<AgentOutcome> {
  return runGuarded(
    a,
    "Word",
    async () => {
      const ooxml = await Word.run(async (ctx) => {
        const xml = ctx.document.body.getOoxml();
        await ctx.sync();
        return xml.value;
      });
      return () =>
        Word.run(async (ctx) => {
          ctx.document.body.insertOoxml(ooxml, "Replace");
          await ctx.sync();
        });
    },
    (fn) =>
      Word.run(async (ctx) => {
        const r = await fn(ctx, Word);
        await ctx.sync();
        return r;
      }),
  );
}

/**
 * PowerPoint (#23): øjebliksbilledet er alle slides eksporteret hver for sig (PowerPointApi 1.8).
 * Ved fejl og fortryd lægges de tilbage i rækkefølge, og de slides, der var der efter scriptet,
 * fjernes. Uden PowerPointApi 1.8 kan der ikke tages et øjebliksbillede, og scriptet køres ikke.
 */
export function runPowerPointScript(a: RunScript): Promise<AgentOutcome> {
  return runGuarded(
    a,
    "PowerPoint",
    async () => {
      if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8"))
        throw new Error("Kodetilstanden i PowerPoint kræver PowerPointApi 1.8 (Microsoft 365), så ændringerne kan fortrydes.");
      const saved = await PowerPoint.run(async (ctx) => {
        const col = ctx.presentation.slides;
        col.load("items/id");
        await ctx.sync();
        const exported = col.items.map((sl) => sl.exportAsBase64());
        await ctx.sync();
        return exported.map((r) => r.value);
      });
      return () =>
        PowerPoint.run(async (ctx) => {
          const cur = ctx.presentation.slides;
          cur.load("items/id");
          await ctx.sync();
          // De slides, der er der nu (efter scriptet), noteres, før de gemte lægges ind.
          const stale = [...cur.items];
          const now = stale.map((x) => x.id);
          // De gemte slides lægges ind efter de nuværende i den oprindelige rækkefølge …
          const known = new Set(now);
          let target = now[now.length - 1] as string | undefined;
          for (const b64 of saved) {
            ctx.presentation.insertSlidesFromBase64(b64, {
              formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
              ...(target ? { targetSlideId: target } : {}),
            });
            await ctx.sync();
            const after = ctx.presentation.slides;
            after.load("items/id");
            await ctx.sync();
            target = after.items.find((x) => !known.has(x.id))?.id ?? target;
            if (target) known.add(target);
          }
          // … og dem, der var der efter scriptet, fjernes (der skal altid være mindst ét slide undervejs).
          for (const sl of stale) sl.delete();
          await ctx.sync();
        });
    },
    (fn) =>
      PowerPoint.run(async (ctx) => {
        const r = await fn(ctx, PowerPoint);
        await ctx.sync();
        return r;
      }),
  );
}
