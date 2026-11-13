import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve, sep } from "node:path";
import ExcelJS from "exceljs";
import { EXPORT_FOLDERS, EXPORT_FILE_NAME, EXPORT_SHEET_NAME, MAX_EXPORT_CELLS, type ExportSheetRequest } from "@miai/shared";

/**
 * Eksport af et ark som .xlsx-fil (tillæg ADR-0012). Bridgen bygger filen med exceljs ud fra
 * det, add-in'et har læst: værdier, formler (med deres resultat), talformater og kolonnebredder.
 * Sikkerhed: mappen kommer fra et fast enum (aldrig fra brugeren), filnavnet valideres mod
 * EXPORT_FILE_NAME, der er ingen overskrivning (409), og alt over 50.000 celler afvises.
 */

/** Hjemmemappen – kan overlades i tests (MIAI_EXPORT_HOME), så der aldrig skrives på den rigge maskinprofil. */
const home = () => process.env.MIAI_EXPORT_HOME ?? homedir();

/** Mappens navn under hjemmemappen pr. enum-værdi (Windows-navne). */
const FOLDER_DIRS: Record<(typeof EXPORT_FOLDERS)[number], string> = {
  desktop: "Desktop",
  documents: "Documents",
  downloads: "Downloads",
};

/**
 * Det fulde stienavn for eksporten. Kaster ved ugyldigt filnavn: navnet må ikke indeholde
 * stipelementer (se EXPORT_FILE_NAME), og stien tjekkes alligevel for at ligge i mappen.
 */
export function exportPath(fileName: string, folder: (typeof EXPORT_FOLDERS)[number]): string {
  if (!EXPORT_FILE_NAME.test(fileName)) throw new Error("Ugyldigt filnavn (brug fx “Svar 2026.xlsx”)");
  const base = resolve(join(home(), FOLDER_DIRS[folder]));
  const file = resolve(base, fileName);
  if (!file.startsWith(base + sep)) throw new Error("Ugyldigt filnavn");
  return file;
}

/** Bygger .xlsx-filen og svarer med dens indhold (Buffer). Kaster ved inkonsekvent data. */
export async function buildXlsx(req: ExportSheetRequest): Promise<Buffer> {
  if (!EXPORT_SHEET_NAME.test(req.name)) throw new Error("Ugyldigt arknavn (maks. 31 tegn, ingen []:*?/\\)");
  const rows = req.values.length;
  if (req.formulas.length !== rows || req.numberFormats.length !== rows) throw new Error("Inkonsekvent data");
  const cols = req.values[0].length;
  if (req.values.some((r, i) => r.length !== cols || req.formulas[i].length !== cols || req.numberFormats[i].length !== cols)) {
    throw new Error("Inkonsekvent data");
  }
  if (rows * cols > MAX_EXPORT_CELLS) throw new Error(`For mange celler (maks. ${MAX_EXPORT_CELLS.toLocaleString("da-DK")})`);

  const [from] = req.address.replace(/\$/g, "").toUpperCase().split(":");
  const col = (s: string) => [...s.replace(/\d+/g, "")].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const row = (s: string) => Number(s.replace(/[A-Z]+/g, ""));
  const c0 = col(from) - 1;
  const r0 = row(from) - 1;

  const wb = new ExcelJS.Workbook();
  wb.creator = "MIAI";
  const ws = wb.addWorksheet(req.name);
  for (let r = 0; r < rows; r++) {
    const line = ws.getRow(r0 + 1 + r);
    for (let c = 0; c < cols; c++) {
      const cell = line.getCell(c0 + 1 + c);
      const formula = req.formulas[r][c];
      // Formel med sit nuværende resultat, så filen viser tal før den åbnes i Excel.
      const v = req.values[r][c];
      cell.value = formula ? { formula: formula.replace(/^=/, ""), result: v === null || v === undefined ? undefined : v } : v;
      const fmt = req.numberFormats[r][c];
      if (fmt && fmt !== "General") cell.numFmt = fmt;
    }
  }
  req.columnWidths.forEach((w, i) => {
    ws.getColumn(c0 + 1 + i).width = w;
  });
  return (await wb.xlsx.writeBuffer()) as unknown as Buffer;
}

/** Filen må ikke eksistere: agenten overskriver aldrig en brugerfil. */
export const fileExists = (path: string) => existsSync(path);
