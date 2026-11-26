import type { AgentAction, DocContext } from "@miai/shared";
import { splitAddress } from "../excelAddress";
import { countOverwrite, fillPieces, writeTarget } from "./excelSheet";
import type { Overwrite } from "./types";

/**
 * Agentens ændringer i Excel (#75): advarsel før overskrivning, ét trin i Excels egen fortryd og
 * midlertidig markering af de ændrede celler. Office.js-delen ligger her; den rene logik (tælling,
 * mål og stykker af fyld) ligger i excelSheet.ts.
 */

/** ExcelApi 1.20 (Windows 2509, Mac 16.100; ikke LTSC): en hel Excel.run kan blive ét trin i Excels fortryd. */
export const canMergeUndo = (): boolean => Office.context.requirements.isSetSupported("ExcelApi", "1.20");

/**
 * Kører batchen, så alle dens context.sync() bliver ét trin i Excels egen fortryd (Ctrl+Z), når
 * ExcelApi 1.20 findes (`mergeUndoGroup`). Ellers som en almindelig Excel.run.
 */
export function runAsOneUndo<T>(batch: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  return canMergeUndo() ? Excel.run({ mergeUndoGroup: true }, batch) : Excel.run(batch);
}

/**
 * Tæller for hver handling, hvor mange celler med indhold og hvor mange formler den vil overskrive
 * i arket, som det er nu. Der læses kun. Et ark, der ikke findes endnu (planen opretter det), har
 * intet at overskrive.
 */
export async function countOverwrites(actions: AgentAction[], dc: DocContext): Promise<(Overwrite | undefined)[]> {
  const active = dc.address?.includes("!") ? splitAddress(dc.address).sheet : "";
  const targets = actions.map((a) => {
    const t = writeTarget(a);
    if (!t || !("address" in a)) return undefined;
    return { ...t, sheet: a.address.includes("!") ? splitAddress(a.address).sheet : active };
  });
  if (!targets.some(Boolean)) return actions.map(() => undefined);
  return Excel.run(async (ctx) => {
    const sheets = new Map<string, Excel.Worksheet>();
    for (const t of targets) {
      if (!t || sheets.has(t.sheet)) continue;
      const ws = t.sheet ? ctx.workbook.worksheets.getItemOrNullObject(t.sheet) : ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      sheets.set(t.sheet, ws);
    }
    await ctx.sync();
    const ranges = targets.map((t) => {
      const ws = t ? sheets.get(t.sheet) : undefined;
      if (!t || !ws || ws.isNullObject) return undefined;
      const r = ws.getRange(t.address);
      r.load("formulas");
      return r;
    });
    await ctx.sync();
    return ranges.map((r, i) => {
      if (!r) return undefined;
      const n = countOverwrite(r.formulas as unknown[][], targets[i]!.incoming);
      return n.cells > 0 ? n : undefined;
    });
  });
}

/** Markeringens fyldfarve: lys gul, som Excels egen "Guld, Accent 4, lysere 80 %". */
export const MARK_COLOR = "#FFF2CC";

/** Et markeret område og cellernes fyldfarve fra før markeringen (undefined = intet fyld). */
export interface CellMark {
  sheet: string;
  /** Uden arknavn. */
  address: string;
  fills: (string | undefined)[][];
}

/** Fyldfarven og mønstret: mønstret "None" er en celle uden fyld (farven læses da som hvid). */
const FILL_PROPS = { format: { fill: { color: true, pattern: true } } } as Excel.CellPropertiesLoadOptions;

/**
 * Markerer områderne med en lys fyldfarve (ExcelApi 1.9) og returnerer det, der skal til for at
 * fjerne markeringen igen. Cellernes fyld læses for alle områder, før noget farves, så overlappende
 * områder ikke gemmer markeringsfarven som "det oprindelige".
 */
export async function markCells(ctx: Excel.RequestContext, sheet: Excel.Worksheet, sheetName: string, addresses: string[]): Promise<CellMark[]> {
  const ranges = [...new Set(addresses)].map((address) => {
    const r = sheet.getRange(address);
    return { address, r, props: r.getCellProperties(FILL_PROPS) };
  });
  if (!ranges.length) return [];
  await ctx.sync();
  const marks = ranges.map((x) => ({
    sheet: sheetName,
    address: x.address,
    fills: x.props.value.map((row) =>
      row.map((p) => {
        const fill = p.format?.fill;
        return fill?.color && fill.pattern !== "None" ? fill.color : undefined;
      }),
    ),
  }));
  for (const x of ranges) x.r.format.fill.color = MARK_COLOR;
  // Afviser Excel farvningen, returneres det gemte fyld alligevel, så en halvt sat markering kan fjernes.
  await ctx.sync().catch(() => {});
  return marks;
}

/**
 * Fjerner markeringen: celler, der havde fyld, får farven tilbage, og celler uden fyld får fyldet
 * ryddet (så fx en tabeltypografis striber og arkets gitterlinjer vises igen). Resten af cellernes
 * format er aldrig rørt af markeringen.
 */
export function unmarkCells(marks: CellMark[]): Promise<void> {
  return Excel.run(async (ctx) => {
    for (const m of marks) {
      const ws = ctx.workbook.worksheets.getItem(m.sheet);
      for (const piece of fillPieces(m.address, m.fills)) {
        const r = ws.getRange(piece.address);
        if (piece.colors) r.setCellProperties(piece.colors.map((row) => row.map((color) => ({ format: { fill: { color } } }))));
        else r.format.fill.clear();
      }
    }
    await ctx.sync();
  });
}
