import { z } from "zod";

/**
 * Eksport af et Excel-ark som .xlsx-fil (tillæg til ADR-0012): add-in'et læser arket
 * (værdier, formler, talformater og kolonnebredder) som JSON og sender det til bridgen,
 * der bygger filen med exceljs. Kun bridge-tilstand – direkte tilstand har ingen eksekverer,
 * der kan skrive filer. Slået fra som standard og kan forbydes af politikken, så organisationen
 * selv styrer, om agenten må skrive filer på maskinen.
 */

/**
 * Filnavn til saveSheet: skal ende i .xlsx og må ikke indeholde stipelementer eller
 * OS-reserverede tegn. Det første tegn må ikke være punkt eller mellemrum (skjulte filer,
 * og Windows rydder trailing dots). Maks. 255 tegn i alt (NTFS-grænsen).
 */
export const EXPORT_FILE_NAME = /^[^\s/\\:*?"<>|.][^/\\:*?"<>|\n\r]{0,249}\.xlsx$/i;

/** Arknavn i den nye fil: Excels regler (maks. 31 tegn, ingen []:*?/\). */
export const EXPORT_SHEET_NAME = /^[^[\]:*?/\\]{1,31}$/;

/** Maks. antal celler, én eksport må indeholde (beskytter bridgen og netværket). */
export const MAX_EXPORT_CELLS = 50_000;

/** Mappene under brugerens hjem, som eksporten må skrives i. */
export const EXPORT_FOLDERS = ["desktop", "documents", "downloads"] as const;

/** Det, add-in'et sender til bridgens POST /api/export. */
export const ExportSheetRequest = z.object({
  /** Filnavn med .xlsx (valideres mod EXPORT_FILE_NAME). */
  fileName: z.string().min(1).max(255),
  /** Hvilken mappe under brugerens hjem (standard: desktop). */
  folder: z.enum(EXPORT_FOLDERS).default("desktop"),
  /** Arkets navn i den nye fil (valideres mod EXPORT_SHEET_NAME). */
  name: z.string().min(1).max(31),
  /** Det used range, der læses (fx "A1:F20"), uden arknavn. */
  address: z.string().min(2).max(40),
  /** Cellernes værdier (""/null = tom celle). */
  values: z.array(z.array(z.union([z.string().max(1_000), z.number(), z.boolean()]).nullable())).min(1).max(5_000),
  /** Formler ("" = ingen formel i cellen, ellers fx "=SUM(B2:B5)"). */
  formulas: z.array(z.array(z.string().max(1_000))).min(1).max(5_000),
  /** Talformat pr. celle ("" = standard). */
  numberFormats: z.array(z.array(z.string().max(80))).min(1).max(5_000),
  /** Kolonnebredder i tegn (Office.js' columnWidth); tom = default. */
  columnWidths: z.array(z.number().min(1).max(500)).max(256).default([]),
});
export type ExportSheetRequest = z.infer<typeof ExportSheetRequest>;
