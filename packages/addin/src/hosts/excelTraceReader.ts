import type { TraceKind } from "@miai/shared";
import {
  areaBox,
  MAX_AREA_READ,
  parseAreas,
  renderTrace,
  TRACE_LIMITS,
  traceGraph,
  type AreaCells,
  type CellRef,
  type TraceArea,
  type TraceLimits,
  type TraceReader,
} from "./excelTrace";

/**
 * Formelsporing (#76), Office.js-delen: spørger Excel om en celles præcedenter og afhængige og
 * læser cellerne. Der læses kun. Gennemløb, beskæring og teksten til modellen er ren logik i
 * excelTrace.ts.
 *
 * API'erne (bekræftet i @types/office-js): Range.getDirectPrecedents() (ExcelApi 1.12),
 * getDirectDependents() (1.13), getPrecedents() (1.14) og getDependents() (1.15) returnerer alle
 * en WorkbookRangeAreas på tværs af ark og kaster ItemNotFound ved context.sync(), når der ingen
 * er. Grafen gennemløbes niveau for niveau med de direkte kald, så dybden kendes, og så det virker
 * fra 1.12; kaldene for hele grafen bruges kun til at sige, hvad en beskåret graf mangler.
 */

const has = (version: string) => Office.context.requirements.isSetSupported("ExcelApi", version);

/** De retninger, værten kan spore: præcedenter fra ExcelApi 1.12, afhængige fra 1.13. */
export function traceSupport(): TraceKind[] {
  return [...(has("1.12") ? (["precedents"] as const) : []), ...(has("1.13") ? (["dependents"] as const) : [])];
}

/** Excel svarer ItemNotFound, når en celle ingen præcedenter eller afhængige har – det er et svar, ikke en fejl. */
const notFound = (e: unknown) => (e as { code?: string })?.code === "ItemNotFound" || /ItemNotFound/.test(e instanceof Error ? e.message : String(e));

const CELL_PROPS = "formulas,values,valueTypes";

export function officeTraceReader(ctx: Excel.RequestContext, kind: TraceKind, opts: { whole?: boolean } = {}): TraceReader {
  const range = (sheet: string, address: string) => ctx.workbook.worksheets.getItem(sheet).getRange(address);
  const areasOf = (p: Excel.WorkbookRangeAreas, sheet: string) =>
    parseAreas(
      p.areas.items.map((a) => a.address),
      sheet,
    );

  /**
   * Alle cellerne spørges i én rundtur. Én celle uden præcedenter får hele rundturen til at fejle
   * (ItemNotFound), og så er intet fra den indlæst; derfor deles cellerne i to og spørges igen,
   * til den enkelte celle står alene.
   */
  const direct = async (cells: CellRef[]): Promise<(TraceArea[] | undefined)[]> => {
    if (!cells.length) return [];
    const asked = cells.map((c) => {
      const r = range(c.sheet, c.address);
      const p = kind === "precedents" ? r.getDirectPrecedents() : r.getDirectDependents();
      p.areas.load("items/address");
      return p;
    });
    try {
      await ctx.sync();
      return asked.map((p, i) => areasOf(p, cells[i].sheet));
    } catch (e) {
      if (cells.length === 1) return [notFound(e) ? [] : undefined];
      const mid = cells.length >> 1;
      return [...(await direct(cells.slice(0, mid))), ...(await direct(cells.slice(mid)))];
    }
  };

  /**
   * Et almindeligt område læses, som det er. Hele kolonner og rækker ("B:B") og meget store
   * områder læses kun i deres brugte del, og højst MAX_AREA_READ celler af den.
   */
  const readAll = async (areas: TraceArea[]): Promise<AreaCells[]> => {
    const first = areas.map((a) => {
      const r = range(a.sheet, a.range);
      const box = areaBox(a.range);
      if (box && (box.r1 - box.r0 + 1) * (box.c1 - box.c0 + 1) <= MAX_AREA_READ) {
        r.load(CELL_PROPS);
        return { block: r, row: box.r0, col: box.c0 };
      }
      const used = r.getUsedRangeOrNullObject(true);
      used.load("rowIndex,columnIndex,rowCount,columnCount");
      return { used };
    });
    await ctx.sync();
    const second = first.map((f) => {
      if (!f.used || f.used.isNullObject) return undefined;
      const cols = Math.min(f.used.columnCount, MAX_AREA_READ);
      const rows = Math.max(1, Math.min(f.used.rowCount, Math.floor(MAX_AREA_READ / cols)));
      const block = f.used.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      block.load(CELL_PROPS);
      return { block, row: f.used.rowIndex, col: f.used.columnIndex, total: f.used.rowCount * f.used.columnCount };
    });
    if (second.some(Boolean)) await ctx.sync();
    return first.map((f, i): AreaCells => {
      const x = f.block ? f : second[i];
      // Intet brugt i området: ingen celler.
      if (!x?.block) return { row: 0, col: 0, formulas: [], values: [] };
      return {
        row: x.row,
        col: x.col,
        formulas: x.block.formulas as unknown[][],
        values: x.block.values as unknown[][],
        types: x.block.valueTypes as string[][],
        ...("total" in x ? { total: x.total } : {}),
      };
    });
  };
  const read = async (areas: TraceArea[]): Promise<(AreaCells | undefined)[]> => {
    try {
      return await readAll(areas);
    } catch {
      // Et område, der ikke kan læses (fx et ark, der er slettet), må ikke koste de andre: ét ad gangen.
      const out: (AreaCells | undefined)[] = [];
      for (const a of areas)
        out.push(
          await readAll([a]).then(
            (r) => r[0],
            () => undefined,
          ),
        );
      return out;
    }
  };

  const canWhole = opts.whole !== false && has(kind === "precedents" ? "1.14" : "1.15");
  const whole = async (cell: CellRef): Promise<TraceArea[] | undefined> => {
    const r = range(cell.sheet, cell.address);
    const p = kind === "precedents" ? r.getPrecedents() : r.getDependents();
    p.areas.load("items/address");
    await ctx.sync();
    return areasOf(p, cell.sheet);
  };
  return { direct, read, ...(canWhole ? { whole } : {}) };
}

/** Sporer fra en celle og returnerer den kompakte tekst til modellen. Ændrer intet i arket. */
export async function traceText(
  ctx: Excel.RequestContext,
  kind: TraceKind,
  start: CellRef,
  limits: TraceLimits = TRACE_LIMITS,
  opts: { whole?: boolean } = {},
): Promise<string> {
  return renderTrace(await traceGraph(start, kind, officeTraceReader(ctx, kind, opts), limits), limits);
}
