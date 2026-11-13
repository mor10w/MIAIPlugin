import { z } from "zod";
import { EXPORT_FILE_NAME } from "./export.js";
import { checkScript } from "./script.js";
import type { DocContext, HostKind } from "./types.js";

/**
 * Agent-tilstand (#6, ADR-0012): modellen foreslår en plan af handlinger fra en fast whitelist.
 * Intet udføres af motoren. Add-in'et viser planen, brugeren vælger og bekræfter, og først
 * derefter udføres handlingerne med Office.js – med fortryd for hele sessionen.
 *
 * Bevidst ingen netværks- eller filværktøjer og ingen fri formatering (skrifttype, farve,
 * størrelse): dokumentindhold kan indeholde prompt injection, og det værste en kapret plan kan
 * gøre, er at foreslå ændringer, som brugeren ser, før de udføres. Formatering sker kun med
 * dokumentets egne typografier og begrænset Markdown (fed, kursiv, links, lister) – i Excel med
 * de indbyggede tabeltypografier, talformater og kolonnebredde. Den ene dokumenterede undtagelse
 * er saveSheet, der gemmer ét ark som .xlsx i en fast mappe (tillæg ADR-0012): kun med bridge,
 * bekræftet som alle andre handlinger, og aldrig destruktiv.
 *
 * Formatering bevares: modellen ser afsnittene med typografi (¶-numre), strukturændringer
 * peger på afsnit (ikke tekstsøgning), og erstat ændrer kun de ord, der er forskellige.
 */

/** Maks. antal handlinger i én plan (loft pr. session). */
export const MAX_AGENT_ACTIONS = 20;
/** Maks. antal celler, én Excel-handling må skrive. */
export const MAX_AGENT_CELLS = 2_000;

const reason = z.string().max(200).describe("Very short reason shown to the user, in the user's language");
const find = z.string().min(1).max(255).describe("Exact text copied verbatim from ONE paragraph of the document (a few words up to one sentence)");
const paragraph = z.number().int().min(1).max(100_000).describe("Paragraph number (¶) from the document structure");

/**
 * Excels indbyggede tabeltypografier. De følger projektmappens tema (farver og skrift), så agenten
 * kan give en tabel et moderne udseende uden fri formatering.
 */
export const TableStyle = z
  .string()
  .regex(/^TableStyle(?:Light(?:[1-9]|1\d|2[01])|Medium(?:[1-9]|1\d|2[0-8])|Dark(?:[1-9]|1[01]))$/)
  .describe("Built-in Excel table style, e.g. TableStyleMedium2 (banded rows, accent header – Excel's default) or TableStyleLight9 (light, accent header)");

/** Et gyldigt arknavn i Excel: højst 31 tegn, uden : \\ / ? * [ ] og uden apostrof først eller sidst. */
const SheetName = z
  .string()
  .min(1)
  .max(31)
  .regex(/^(?!')[^:\\/?*[\]]+(?<!')$/);

/** Indbyggede typografier, agenten må bruge (ingen fri formatering: skrifttype, farve og størrelse følger dokumentets typografier). */
export const WordStyle = z.enum(["normal", "heading1", "heading2", "heading3", "bullet", "numbered", "quote"]);
export type WordStyle = z.infer<typeof WordStyle>;

export const AgentAction = z.discriminatedUnion("action", [
  // --- Word ------------------------------------------------------------------
  z.object({
    action: z.literal("replace"),
    find,
    replace: z.string().max(4_000).describe("Plain text. Only the words that differ are changed, so the rest keeps its formatting"),
    all: z.boolean().optional().describe("Replace every occurrence (default: only the first)"),
    reason,
  }),
  z.object({
    action: z.literal("insert"),
    text: z
      .string()
      .min(1)
      .max(6_000)
      .describe(
        "New paragraph(s), one per line. Limited Markdown: '# '/'## '/'### ' headings, '- ' bullets, '1. ' numbered, '> ' quote, **bold**, *italic*, [text](https://link)",
      ),
    position: z.enum(["before", "after", "start", "end"]).describe("before/after paragraph ¶, or at the start/end of the document"),
    paragraph: paragraph.optional().describe("Required for before/after"),
    style: WordStyle.optional().describe("Style for lines without a Markdown prefix (default: normal)"),
    reason,
  }),
  z.object({
    action: z.literal("delete"),
    paragraph: paragraph.optional().describe("Delete this whole paragraph (preferred for whole paragraphs)"),
    find: find.optional().describe("Or: delete only this text inside a paragraph"),
    reason,
  }),
  z.object({
    action: z.literal("move"),
    paragraph: paragraph.describe("Paragraph to move (with its formatting)"),
    to: paragraph.describe("Target paragraph"),
    position: z.enum(["before", "after"]),
    reason,
  }),
  z.object({ action: z.literal("setStyle"), paragraph, style: WordStyle, reason }),
  z.object({ action: z.literal("comment"), find, text: z.string().min(1).max(2_000), reason }),
  z.object({
    action: z.literal("insertTable"),
    position: z.enum(["before", "after", "start", "end"]).describe("before/after: relative to ¶paragraph; start/end: of the document or selection"),
    paragraph: paragraph.optional().describe("Required for before/after – a paragraph that is NOT in a table"),
    rows: z
      .array(z.array(z.string().max(500)).min(1).max(12))
      .min(2)
      .max(60)
      .describe("The table's cells, row by row. The FIRST row is the header row. Plain text, one string per cell"),
    reason,
  }),
  z.object({
    action: z.literal("addTableRow"),
    paragraph: paragraph.describe("A paragraph (¶) in the table row next to which the new row goes"),
    position: z.enum(["before", "after"]).default("after"),
    values: z.array(z.string().max(500)).min(1).max(12).describe("The new row's cells from left to right"),
    reason,
  }),
  z.object({ action: z.literal("deleteTableRow"), paragraph: paragraph.describe("A paragraph (¶) in the table row to delete"), reason }),
  // --- PowerPoint (replace deles med Word) --------------------------------------
  z.object({
    action: z.literal("setText"),
    shape: z
      .string()
      .regex(/^\d{1,3}\.\d{1,3}$/)
      .describe("Shape reference 'slide.shape' from the structure, e.g. '2.1'"),
    text: z
      .string()
      .max(6_000)
      .describe("New text, one line per paragraph ('- ' prefixes are optional: the shape's own bullet style is used); **bold**, *italic*"),
    reason,
  }),
  z.object({
    action: z.literal("addSlide"),
    after: z.number().int().min(0).max(1_000).describe("Insert after this slide number (0 = first)"),
    layout: z.string().max(120).optional().describe("Layout name from the presentation's layouts (default: the layout of the slide before)"),
    title: z.string().max(500),
    body: z.string().max(6_000).optional().describe("Body text, one line per bullet"),
    reason,
  }),
  z.object({ action: z.literal("deleteSlide"), slide: z.number().int().min(1).max(1_000), reason }),
  z.object({
    action: z.literal("moveSlide"),
    slide: z.number().int().min(1).max(1_000),
    to: z.number().int().min(1).max(1_000).describe("New position (slide number)"),
    reason,
  }),
  // --- Excel -----------------------------------------------------------------
  z.object({
    action: z.literal("setValues"),
    address: z.string().min(2).max(40).describe("Top-left cell or range on the active sheet, e.g. 'D2' or 'D2:E5'"),
    values: z
      .array(z.array(z.union([z.string().max(1_000), z.number(), z.boolean(), z.null()])))
      .min(1)
      .max(200)
      .describe("Rows of cell values written from the top-left cell. Numbers as JSON numbers (not text), null = leave empty"),
    reason,
  }),
  z.object({
    action: z.literal("setFormula"),
    address: z.string().min(2).max(20).describe("One cell, e.g. 'F2'"),
    formula: z.string().min(2).max(1_000).describe("Excel formula starting with '=', English function names, comma separator"),
    reason,
  }),
  z.object({ action: z.literal("clear"), address: z.string().min(2).max(40).describe("Range on the active sheet whose contents are cleared"), reason }),
  z.object({
    action: z.literal("appendRows"),
    table: z.string().max(255).describe("Excel table name from the structure"),
    values: z
      .array(z.array(z.union([z.string().max(1_000), z.number(), z.boolean(), z.null()])))
      .min(1)
      .max(200)
      .describe("Rows to add at the end of the table, one value per column (null for calculated columns)"),
    reason,
  }),
  z.object({
    action: z.literal("deleteRows"),
    address: z
      .string()
      .min(2)
      .max(40)
      .describe(
        "The rows to delete as a range across the data's full width, e.g. 'A5:D7'. Rows inside an Excel table are deleted as table rows; cells below move up",
      ),
    reason,
  }),
  z.object({
    action: z.literal("deleteEmptyRows"),
    target: z
      .string()
      .min(1)
      .max(255)
      .describe("Excel table name from the structure, or a range such as 'A1:F200'. Every completely empty row in it is deleted"),
    reason,
  }),
  z.object({
    action: z.literal("formatTable"),
    target: z
      .string()
      .min(1)
      .max(255)
      .describe("Existing Excel table name, or the data range INCLUDING its header row (e.g. 'A1:F20') to turn into a formatted Excel table"),
    style: TableStyle,
    reason,
  }),
  z.object({
    action: z.literal("setNumberFormat"),
    address: z.string().min(2).max(40).describe("Range on the active sheet, e.g. 'C2:C40'"),
    format: z.string().min(1).max(80).describe('Excel number format code, e.g. "#,##0", "#,##0.00 kr.", "0%", "dd-mm-yyyy"'),
    reason,
  }),
  z.object({
    action: z.literal("sortRange"),
    target: z.string().min(1).max(255).describe("Excel table name, or the data range INCLUDING its header row (e.g. 'A1:F40')"),
    by: z.string().min(1).max(255).describe("Column to sort by: header text (e.g. 'Dato') or column letter (e.g. 'C')"),
    order: z.enum(["asc", "desc"]).default("asc"),
    reason,
  }),
  z.object({
    action: z.literal("conditionalFormat"),
    address: z.string().min(2).max(40).describe("Range of values, without the header, e.g. 'C2:C40'"),
    rule: z
      .enum(["colorScale", "dataBar", "greaterThan", "lessThan", "duplicates", "top10"])
      .describe(
        "colorScale = red-yellow-green scale; dataBar = bars in the cells; greaterThan/lessThan = highlight cells compared with value; duplicates; top10 = highlight the 10 largest",
      ),
    value: z.number().optional().describe("Only for greaterThan/lessThan"),
    reason,
  }),
  z.object({
    action: z.literal("addChart"),
    source: z.string().min(1).max(255).describe("Excel table name or data range including headers, e.g. 'A1:B13'"),
    type: z.enum(["column", "bar", "line", "pie", "area", "scatter"]),
    title: z.string().max(120).optional(),
    reason,
  }),
  z.object({
    action: z.literal("readRange"),
    address: z.string().min(2).max(300).describe("Range to look at before planning, with sheet name when it is on another sheet, e.g. 'Budget!A1:F40'"),
    reason,
  }),
  z.object({ action: z.literal("addSheet"), name: SheetName.describe("Name of the new, empty sheet (max 31 characters)"), reason }),
  z.object({
    action: z.literal("renameSheet"),
    sheet: z.string().min(1).max(255).describe("Current name of the sheet"),
    name: SheetName.describe("New name (max 31 characters)"),
    reason,
  }),
  z.object({
    action: z.literal("addPivotTable"),
    source: z.string().min(1).max(255).describe("Excel table name, or the data range INCLUDING its header row (e.g. 'Salg!A1:D200')"),
    rows: z.array(z.string().min(1).max(255)).min(1).max(4).describe("Header texts of the fields to group by as rows, e.g. ['Region']"),
    columns: z.array(z.string().min(1).max(255)).max(3).default([]).describe("Header texts of fields to spread across columns (often empty)"),
    values: z
      .array(
        z.object({
          field: z.string().min(1).max(255).describe("Header text of the field to summarise"),
          summarize: z.enum(["sum", "count", "average", "max", "min"]).default("sum"),
        }),
      )
      .min(1)
      .max(4),
    destination: z.string().min(2).max(300).describe("Top-left cell for the pivot table, in an empty area or on a new sheet, e.g. 'F1' or 'Pivot!A1'"),
    reason,
  }),
  z.object({
    action: z.literal("setValidation"),
    address: z.string().min(2).max(300).describe("Cells that get the rule, without the header, e.g. 'D2:D200'"),
    list: z.array(z.string().min(1).max(100)).min(1).max(50).optional().describe("Allowed values shown as a drop-down list, e.g. ['Ny', 'I gang', 'Færdig']"),
    min: z.number().optional().describe("Smallest allowed number (when no list)"),
    max: z.number().optional().describe("Largest allowed number (when no list)"),
    whole: z.boolean().optional().describe("true = only whole numbers"),
    reason,
  }),
  z.object({
    action: z.literal("freezeRows"),
    rows: z.number().int().min(0).max(50).describe("Number of top rows to keep visible when scrolling (1 = the header row); 0 removes freezing"),
    reason,
  }),
  z.object({
    action: z.literal("runScript"),
    description: z.string().min(1).max(300).describe("What the script does, in the user's language – shown to the user before it runs"),
    code: z.string().min(1).max(8_000).describe("Body of an async function with `context` (Excel.RequestContext) and `Excel` in scope"),
    reason,
  }),
  z.object({
    action: z.literal("autofitColumns"),
    address: z.string().min(2).max(40).describe("Range whose columns get a width that fits their content, e.g. 'A1:F40'"),
    reason,
  }),
  z.object({
    action: z.literal("saveSheet"),
    fileName: z
      .string()
      .min(1)
      .max(255)
      .describe('File name with the .xlsx extension, e.g. "Svar 2026.xlsx" – no folders, spaces at the start or special characters'),
    sheet: z.string().min(1).max(31).optional().describe("Name of the sheet to save (default: the active sheet)"),
    reason,
  }),
]);
export type AgentAction = z.infer<typeof AgentAction>;
export type AgentActionKind = AgentAction["action"];

export const AgentPlan = z.object({
  summary: z.string().max(1_000).describe("One or two sentences: what the plan does, or why it cannot be done with the available actions"),
  actions: z.array(AgentAction).max(MAX_AGENT_ACTIONS),
});
export type AgentPlan = z.infer<typeof AgentPlan>;

export const WORD_ACTIONS: AgentActionKind[] = ["replace", "insert", "delete", "move", "setStyle", "comment", "insertTable", "addTableRow", "deleteTableRow"];
export const EXCEL_ACTIONS: AgentActionKind[] = [
  "setValues",
  "setFormula",
  "appendRows",
  "clear",
  "deleteRows",
  "deleteEmptyRows",
  "formatTable",
  "setNumberFormat",
  "autofitColumns",
  "sortRange",
  "conditionalFormat",
  "addChart",
  "freezeRows",
  "addPivotTable",
  "setValidation",
  "addSheet",
  "renameSheet",
  "readRange",
  "runScript",
  "saveSheet",
];
export const POWERPOINT_ACTIONS: AgentActionKind[] = ["replace", "setText", "addSlide", "deleteSlide", "moveSlide"];
export const agentActionsFor = (host: HostKind): AgentActionKind[] =>
  host === "word" ? WORD_ACTIONS : host === "excel" ? EXCEL_ACTIONS : host === "powerpoint" ? POWERPOINT_ACTIONS : [];

/** Destruktive handlinger fremhæves i planen. */
export const isDestructive = (a: AgentAction) =>
  a.action === "delete" ||
  a.action === "deleteTableRow" ||
  a.action === "clear" ||
  a.action === "deleteSlide" ||
  a.action === "deleteRows" ||
  a.action === "deleteEmptyRows" ||
  a.action === "runScript" ||
  (a.action === "replace" && !a.replace) ||
  (a.action === "setText" && !a.text.trim());

/** PowerPoint-strukturen til prompten: slides med layout og figurer som "2.1 [body] …". */
export function renderSlides(slides: NonNullable<DocContext["slides"]>, layouts?: string[]): string {
  const lines = slides.flatMap((sl, i) => [
    `Slide ${i + 1} (layout: ${sl.layout})`,
    ...sl.shapes.map((sh, j) => `  ${i + 1}.${j + 1} [${sh.kind}] ${sh.name}: ${sh.text.replace(/\n/g, " ⏎ ")}`),
  ]);
  if (layouts?.length) lines.push("", `Layouts: ${layouts.join(" | ")}`);
  return lines.join("\n");
}

/** Excel-strukturen til prompten: tabeller og kolonnernes talformat. */
export function renderExcelStructure(x: NonNullable<DocContext["excel"]>): string {
  const lines: string[] = [];
  if (x.sheets?.length)
    lines.push(
      "Sheets:",
      ...x.sheets.map(
        (sh) =>
          `- ${sh.name}${sh.name === x.activeSheet ? " (active)" : ""}: ${sh.used ? `${sh.used}${sh.headers.length ? `, headers: ${sh.headers.join(" | ")}` : ""}` : "empty"}`,
      ),
    );
  if (x.tables.length) lines.push("Tables:", ...x.tables.map((t) => `- ${t.name} (${t.address}): ${t.headers.join(" | ")}`));
  if (x.columnFormats.length) lines.push(`Number formats: ${x.columnFormats.join("; ")}`);
  return lines.join("\n");
}

/**
 * Dokumentets afsnit som tekst til prompten: "¶3 [Heading2] Priser", med (list) og (table).
 * Lange afsnit vises i fuld længde, så modellen kan citere ordret.
 */
export function renderParagraphs(paras: NonNullable<DocContext["paragraphs"]>): string {
  // Tabelceller vises med tabel, række og kolonne (#24), så agenten kan tilføje og slette rækker.
  const cell = (p: (typeof paras)[number]) => (p.cell ? ` (table ${p.cell.table}, row ${p.cell.row}, col ${p.cell.col})` : p.table ? " (table)" : "");
  return paras.map((p, i) => `¶${i + 1} [${p.style}]${p.list ? " (list)" : ""}${cell(p)} ${p.text}`).join("\n");
}

/** Beskrivelse af værktøjerne til systemprompten – kun dem, der gælder for værten. */
/**
 * Kodetilstand (ADR-0017): beskrivelsen af runScript til prompten – kun når den er slået til.
 * Reglerne svarer til sandkassens (checkScript): ingen globale objekter ud over de sikre, ingen
 * this/klasser/import, og egenskaber, der starter med _, er spærret.
 */
export const SCRIPT_GUIDE = [
  "- runScript {description, code}: ADVANCED, only when none of the actions above can do it (e.g. named ranges, sorting by several columns, chart settings, pivot table filters). It must be the ONLY action in the plan.",
  "  `code` is the body of an async function with `context` (Excel.RequestContext) and `Excel` in scope. Use Office.js (ExcelApi): `const sheet = context.workbook.worksheets.getActiveWorksheet();`, call `.load(...)` and `await context.sync()` before reading properties, and finish with `await context.sync()`.",
  "  Only these globals exist: Math, JSON, Number, String, Boolean, Array, Object, Date, Map, Set, Promise, Error, isNaN, parseInt, parseFloat, console. No fetch, window, document, eval, this, classes or import; property names starting with _ are blocked.",
  "  Write complete, working code (no placeholders). Optionally `return` a short text describing what changed – it is shown to the user. Never change other sheets than the active one unless asked.",
].join("\n");

/** Eksport (tillæg ADR-0012): beskrivelsen af saveSheet til prompten – kun når den er slået til. */
export const SAVE_GUIDE = [
  "- saveSheet {fileName, sheet?, reason}: save a whole sheet as a NEW .xlsx file on the user's desktop (Documents/Downloads if asked), preserving values, formulas and column widths. It does not change the workbook – use it when the user wants a separate file (e.g. 'put this sheet in its own file' or 'give me this as a new Excel file').",
].join("\n");

/** Maks. antal celler, agenten kan læse pr. readRange (#20). */
export const MAX_READ_CELLS = 6_000;
/** Maks. antal runder, hvor agenten læser, før den skal lave sin plan. */
export const MAX_READ_ROUNDS = 3;

/**
 * Agent i flere trin (#20): beskrivelsen af readRange – kun så længe agenten har læse-runder tilbage.
 * En plan med readRange vises ikke for brugeren: add-in'et henter cellerne (ingen ændringer) og
 * spørger modellen igen med værdierne.
 */
export const READ_GUIDE = [
  "- readRange {address}: LOOK at cells you need but cannot see – another sheet, or a range outside the selection (e.g. 'Budget!A1:F40'). Nothing is changed and the user is not asked: the add-in returns the values and asks you again.",
  "  When you need to read, return ONLY readRange actions (up to 4) and nothing else; make the real plan when you have the values. Do not read what is already shown, and do not read at all when the request can be planned from what you see.",
].join("\n");

export function agentToolGuide(host: HostKind, opts: { scripts?: boolean; export?: boolean; reads?: boolean } = {}): string {
  if (host === "excel")
    return [
      "Available actions (Excel, active sheet only):",
      '- setValues {address, values}: write a 2D array of values starting at address (e.g. "D2"). Use JSON numbers for numbers (1200.5, not "1.200,50"); null leaves a cell empty.',
      '- setFormula {address, formula}: put ONE formula in ONE cell (e.g. "F2"). For a column of formulas, use one action per cell (max 20 actions).',
      "- appendRows {table, values}: add rows at the end of an Excel table (listed under Tables). The rows get the table's formatting and calculated columns – always prefer this over setValues below a table. Use null for calculated columns.",
      "- clear {address}: clear the contents of a range (formatting is kept). Use deleteRows instead when rows should disappear.",
      "- deleteRows {address}: delete whole rows of data (e.g. 'A5:D7'); cells below move up. Inside an Excel table, the table shrinks.",
      "- deleteEmptyRows {target}: delete every completely empty row in a table (by name) or range – use this for 'remove the empty rows' instead of listing rows yourself.",
      "- formatTable {target, style}: give a table a modern, readable look with one of Excel's built-in table styles (they follow the workbook's theme). target = an existing table name, or the data range including its header row, which then becomes an Excel table with filter buttons and banded rows. Good choices: TableStyleMedium2 (default), TableStyleMedium9, TableStyleLight9, TableStyleLight1 (minimal).",
      "- setNumberFormat {address, format}: number format for a range, e.g. '#,##0' (thousands separator), '#,##0.00 kr.', '0%', 'dd-mm-yyyy'.",
      "- autofitColumns {address}: make the columns as wide as their content.",
      "- sortRange {target, by, order}: sort a table (by name) or a data range with its header row by one column (header text or letter), asc or desc.",
      "- conditionalFormat {address, rule, value?}: highlight values with Excel's standard formats – colorScale, dataBar, greaterThan/lessThan (with value), duplicates, top10. Address = the values without the header.",
      "- addChart {source, type, title?}: a chart (column, bar, line, pie, area, scatter) of a table or a range with headers; it is placed next to the data.",
      "- freezeRows {rows}: keep the top rows visible when scrolling (1 = the header row; 0 = remove).",
      "- addPivotTable {source, rows, columns?, values, destination}: a pivot table of a table (by name) or a range with headers. rows/columns/values name fields by their header text; each value has summarize sum (default), count, average, max or min. Put it in an empty area ('F1') or on a new sheet (addSheet Pivot, then destination 'Pivot!A1').",
      "- setValidation {address, list? | min?, max?, whole?}: data validation – a drop-down list of allowed values, or a number between min and max (whole: true for whole numbers).",
      "- addSheet {name}: add a new, empty sheet. renameSheet {sheet, name}: rename a sheet.",
      "Other sheets: every address may carry a sheet name – Salg!A1:C5, or 'Mit ark'!B2 when the name has spaces. Without a sheet name the address is on the active sheet. A sheet added with addSheet can be written to in the same plan (e.g. addSheet Oversigt, then setValues on Oversigt!A1 and setFormula =SUM(Salg!C2:C40) on Oversigt!B2). Tables are found by name on whichever sheet they are. The sheets and their headers are listed under Sheets; you only see the values of the selection.",
      "For 'make the table nice/modern/readable': formatTable, then setNumberFormat for number and date columns where needed, then autofitColumns. Do not change cell values for this.",
      "Formatting is preserved: values never change cell formats, and new cells directly below or next to data get the neighbouring cells' formats. Fonts, colours and borders come only from the table style – there is no action for free formatting.",
      "Use addresses that match the sheet overview; row numbers are shown in the first column of the selected values. Never write over existing data unless the user asked for it. Addresses refer to the sheet before any of your actions – the add-in deletes rows last, from the bottom up.",
      opts.scripts ? SCRIPT_GUIDE : undefined,
      opts.export ? SAVE_GUIDE : undefined,
      opts.reads ? READ_GUIDE : undefined,
    ]
      .filter(Boolean)
      .join("\n");
  if (host === "powerpoint")
    return [
      "Available actions (PowerPoint). The presentation is shown slide by slide with layout and shapes referenced as slide.shape (e.g. 2.1); [title]/[body] are placeholders from the layout, [text] is a free text box.",
      "- replace {find, replace, all?}: change words inside a shape's text. `find` is copied verbatim from one shape; only the words that differ are replaced, so formatting is kept.",
      "- setText {shape, text}: replace all text in one shape. One line per paragraph (no bullet characters needed – the placeholder adds its own); **bold**, *italic*. The shape keeps its own font, size, colour and bullet style.",
      "- addSlide {after, layout?, title, body?}: add a slide from one of the presentation's layouts (listed under Layouts) so it follows the template; the title and body go into the layout's placeholders.",
      "- deleteSlide {slide}, moveSlide {slide, to}: remove or reorder slides.",
      "Preserve the design: prefer replace for small changes, setText only when most of a shape's text changes, and addSlide with an existing layout instead of rebuilding slides. Never add your own text boxes, colours or fonts.",
      "Slide and shape numbers refer to the presentation as shown, before any of your actions.",
    ].join("\n");
  return [
    "Available actions (Word). The document is shown as numbered paragraphs: ¶number [style] text; (list) = list item, (table 1, row 2, col 1) = inside a table cell.",
    "- replace {find, replace, all?}: change text inside a paragraph. `find` is copied verbatim from one paragraph; only the words that differ are replaced, so bold, links and other formatting around them are kept.",
    "- insert {text, position: before|after|start|end, paragraph?, style?}: add new paragraph(s) before/after ¶paragraph or at the start/end. One line per paragraph; use Markdown for structure and emphasis ('## ' heading, '- ' bullet, '1. ' numbered, **bold**, *italic*, [text](https://…)). Lines without a prefix get `style` (default normal) – the new text never inherits the neighbour's formatting.",
    "- delete {paragraph} removes a whole paragraph; delete {find} removes only that text inside a paragraph.",
    "- move {paragraph, to, position: before|after}: move a whole paragraph with its formatting.",
    "- setStyle {paragraph, style}: change a paragraph's role (e.g. body text → heading2, or a line → bullet) without touching its text or inline formatting.",
    "- comment {find, text}: add a review comment on the text.",
    "Styles: normal, heading1, heading2, heading3, bullet, numbered, quote. Use the document's existing structure: match the heading levels and list style already used around the place you edit.",
    "Preserve formatting: prefer small replace actions over rewriting whole paragraphs; use setStyle instead of delete + insert to change a role; never delete, move or insert text inside paragraphs marked (table …) – edit their text with replace, and change a table's rows with addTableRow/deleteTableRow.",
    "- insertTable {position, paragraph?, rows}: add a new table before/after ¶paragraph (not inside a table) or at the start/end. rows = the cells row by row, the FIRST row is the header; every row has the same number of cells. The table gets the document's table style. Use it for tabular content instead of lines with tabs or | characters.",
    "- addTableRow {paragraph, position?, values}: add one row to an existing table, before or after (default) the row that contains ¶paragraph; values = the cells from left to right. deleteTableRow {paragraph}: delete the row that contains ¶paragraph.",
    "Paragraph numbers refer to the document as shown, before any of your actions; do not renumber after earlier actions.",
  ].join("\n");
}

export type AgentProblem =
  | "wrong-host"
  | "not-found"
  | "bad-address"
  | "too-large"
  | "anchor-missing"
  /** Afsnitsnummeret findes ikke. */
  | "bad-paragraph"
  /** Strukturændring i en tabel (ville ødelægge tabellen). */
  | "in-table"
  /** Word: afsnittet står ikke i en tabel (kræves for at tilføje eller slette en tabelrække). */
  | "not-in-table"
  /** Handlingen er ufuldstændig (fx delete uden afsnit og tekst). */
  | "invalid"
  /** Word uden afsnitsstruktur (før WordApi 1.3): afsnitshandlinger kan ikke bruges. */
  | "no-structure"
  /** PowerPoint: slide eller figur findes ikke; layout findes ikke i præsentationen. */
  | "bad-slide"
  | "bad-shape"
  | "bad-layout"
  /** Excel: tabellen findes ikke på arket. */
  | "bad-table"
  /** Excel: arket findes ikke i projektmappen (eller navnet er allerede brugt). */
  | "bad-sheet"
  /** Excel: et felt i en pivottabel findes ikke blandt kildens overskrifter. */
  | "bad-field"
  /** Kodetilstanden er ikke slået til (eller forbudt af politikken). */
  | "scripts-off"
  /** Eksporten er ikke slået til (eller forbudt af politikken), eller der er ingen bridge. */
  | "no-bridge"
  /** Et script skal stå alene i planen (fortryd gendanner arket fra før scriptet). */
  | "with-script"
  /** Scriptet bruger noget, sandkassen ikke tillader (se checkScript). */
  | "unsafe-script";

export interface CheckedAction {
  action: AgentAction;
  /** Udfyldt, hvis handlingen ikke kan udføres (vises i planen, og handlingen fravælges). */
  problem?: AgentProblem;
}

const CELL = /^\$?[A-Z]{1,3}\$?[1-9]\d{0,6}$/;
const RANGE = /^(?:(?:'[^']+'|[^!'\s]+)!)?\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?$/i;

/** "D2:E5" → antal celler (uden arknavn). */
export function cellCount(address: string): number {
  const a = address.split("!").pop()!.replace(/\$/g, "").toUpperCase();
  const [from, to = from] = a.split(":");
  const col = (s: string) => [...s.replace(/\d+/g, "")].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const row = (s: string) => Number(s.replace(/[A-Z]+/g, ""));
  return (Math.abs(col(to) - col(from)) + 1) * (Math.abs(row(to) - row(from)) + 1);
}

/** Maks. antal rækker, én sletning må fjerne. */
export const MAX_AGENT_ROWS = 1_000;

/** "Ark1!B2:D4" → 0-baserede, inklusive række- og kolonneindeks (som Office.js' rowIndex/columnIndex). */
export function rangeBox(address: string): { r0: number; c0: number; r1: number; c1: number } {
  const a = address.split("!").pop()!.replace(/\$/g, "").toUpperCase();
  const [from, to = from] = a.split(":");
  const col = (x: string) => [...x.replace(/\d+/g, "")].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const row = (x: string) => Number(x.replace(/[A-Z]+/g, "")) - 1;
  return { r0: Math.min(row(from), row(to)), c0: Math.min(col(from), col(to)), r1: Math.max(row(from), row(to)), c1: Math.max(col(from), col(to)) };
}

type Box = ReturnType<typeof rangeBox>;
const overlaps = (a: Box, b: Box) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
/** Arknavnet i en adresse ("Salg!A1", "'Mit ark'!B2"), eller undefined hvis adressen gælder det aktive ark. */
export function sheetOfAddress(address: string): string | undefined {
  const i = address.lastIndexOf("!");
  if (i < 0) return undefined;
  return address
    .slice(0, i)
    .replace(/^'(.*)'$/, "$1")
    .replace(/''/g, "'");
}

const tablesOf = (ctx: Pick<DocContext, "excel">) =>
  (ctx.excel?.tables ?? []).map((t) => ({ name: t.name, box: rangeBox(t.address), sheet: sheetOfAddress(t.address) ?? ctx.excel?.activeSheet }));
/** Overskrifterne for et område, når det begynder i arkets første brugte række (ellers ukendt). */
function headersOf(ctx: Pick<DocContext, "excel">, address: string): string[] | undefined {
  const name = (sheetOfAddress(address) ?? ctx.excel?.activeSheet)?.toLowerCase();
  const sheet = ctx.excel?.sheets?.find((sh) => sh.name.toLowerCase() === name);
  if (!sheet?.used || !sheet.headers.length) return undefined;
  const used = rangeBox(sheet.used);
  const b = rangeBox(address);
  if (b.r0 !== used.r0 || b.c0 < used.c0) return undefined;
  const from = b.c0 - used.c0;
  // Overskrifterne er afkortet til de første kolonner; rækker området længere ud, kendes de ikke alle.
  return b.c1 - used.c0 < sheet.headers.length ? sheet.headers.slice(from, b.c1 - used.c0 + 1) : undefined;
}

/** Tabellerne på samme ark som adressen (uden arknavn: det aktive ark). Kendes arket ikke, tæller alle med. */
const tablesOn = (ctx: Pick<DocContext, "excel">, address: string) => {
  const sheet = sheetOfAddress(address) ?? ctx.excel?.activeSheet;
  return tablesOf(ctx).filter((t) => !sheet || !t.sheet || t.sheet === sheet);
};

/** Planens læsninger (#20), som add-in'et kan udføre uden bekræftelse: gyldig adresse, kendt ark og ikke for stor. */
export function readRequests(plan: AgentPlan, ctx: Pick<DocContext, "host" | "excel">): Extract<AgentAction, { action: "readRange" }>[] {
  if (ctx.host !== "excel") return [];
  const sheets = ctx.excel?.sheets?.length ? new Set(ctx.excel.sheets.map((sh) => sh.name.toLowerCase())) : undefined;
  return plan.actions.flatMap((a) => {
    if (a.action !== "readRange" || !RANGE.test(a.address) || cellCount(a.address) > MAX_READ_CELLS) return [];
    const sheet = sheetOfAddress(a.address);
    return sheet && sheets && !sheets.has(sheet.toLowerCase()) ? [] : [a];
  });
}

/**
 * Markerede værdier med arkets rækkenummer først, så modellen kan pege på rækker (deleteRows)
 * uden at tælle selv: "5\tJylland\t120".
 */
export function numberedRows(tsv: string, address?: string): string {
  if (!address || !RANGE.test(address)) return tsv;
  const first = rangeBox(address).r0 + 1;
  return tsv
    .split("\n")
    .map((line, i) => `${first + i}\t${line}`)
    .join("\n");
}

/**
 * Tjekker planen mod værten og dokumentet, før den vises: handlinger til den forkerte vært,
 * tekst, der ikke findes i dokumentet, ugyldige adresser og for store skrivninger markeres.
 */
export function checkPlan(
  plan: AgentPlan,
  ctx: Pick<DocContext, "host" | "text" | "paragraphs" | "slides" | "layouts" | "excel">,
  opts: { scripts?: boolean; export?: boolean } = {},
): CheckedAction[] {
  const allowed = agentActionsFor(ctx.host);
  const hasScript = plan.actions.some((a) => a.action === "runScript");
  // Arknavne sammenlignes uden forskel på store og små bogstaver, som i Excel. Uden arkliste
  // (ældre kontekst) kan arknavne ikke tjekkes her; Excel afviser dem så ved udførelsen.
  const existingSheets = new Set((ctx.excel?.sheets ?? []).map((sh) => sh.name.toLowerCase()));
  const knownSheets = ctx.excel?.sheets?.length
    ? new Set([...existingSheets, ...plan.actions.flatMap((a) => (a.action === "addSheet" ? [a.name.toLowerCase()] : []))])
    : undefined;
  const paras = ctx.paragraphs;
  const slides = ctx.slides;
  const text = paras ? paras.map((p) => p.text).join("\n") : (ctx.text ?? "");
  // Søgning sker inden for ét afsnit (Word) eller én figur (PowerPoint) ad gangen.
  const inDoc = (s?: string) =>
    !!s &&
    !s.includes("\n") &&
    (slides ? slides.some((sl) => sl.shapes.some((sh) => sh.text.includes(s))) : paras ? paras.some((p) => p.text.includes(s)) : text.includes(s));
  const slideOk = (n: number) => !!slides && n >= 1 && n <= slides.length;
  /** Problem med et afsnitsnummer, eller undefined hvis det er i orden. */
  const para = (n: number | undefined, structural: boolean): AgentProblem | undefined => {
    if (n === undefined) return "invalid";
    if (!paras) return "no-structure";
    const p = paras[n - 1];
    if (!p) return "bad-paragraph";
    if (structural && p.table) return "in-table";
    return undefined;
  };
  const ok = (action: AgentAction, problem?: AgentProblem): CheckedAction => (problem ? { action, problem } : { action });
  return plan.actions.slice(0, MAX_AGENT_ACTIONS).map((action): CheckedAction => {
    if (!allowed.includes(action.action)) return { action, problem: "wrong-host" };
    if (hasScript && plan.actions.length > 1 && action.action !== "runScript") return { action, problem: "with-script" };
    // Flere ark (#21): et arknavn i en adresse skal findes i projektmappen eller oprettes af planen.
    if (ctx.host === "excel") {
      const refs = [
        "address" in action ? action.address : "target" in action ? action.target : "source" in action ? action.source : undefined,
        "destination" in action ? action.destination : undefined,
      ];
      for (const ref of refs) {
        const sheet = ref && RANGE.test(ref) ? sheetOfAddress(ref) : undefined;
        if (sheet && knownSheets && !knownSheets.has(sheet.toLowerCase())) return { action, problem: "bad-sheet" };
      }
    }
    switch (action.action) {
      case "readRange":
        // Læsning udføres af add-in'et før planen vises; en læsning, der står tilbage i planen, kan ikke udføres.
        if (!RANGE.test(action.address)) return ok(action, "bad-address");
        return ok(action, cellCount(action.address) > MAX_READ_CELLS ? "too-large" : "invalid");
      case "addSheet":
        return ok(action, knownSheets && existingSheets.has(action.name.toLowerCase()) ? "bad-sheet" : undefined);
      case "renameSheet":
        if (!knownSheets) return ok(action);
        return ok(action, !existingSheets.has(action.sheet.toLowerCase()) || knownSheets.has(action.name.toLowerCase()) ? "bad-sheet" : undefined);
      case "runScript": {
        if (!opts.scripts) return ok(action, "scripts-off");
        return ok(action, checkScript(action.code).ok ? undefined : "unsafe-script");
      }
      case "sortRange":
      case "addChart": {
        const target = action.action === "sortRange" ? action.target : action.source;
        if (tablesOf(ctx).some((t) => t.name === target)) return ok(action);
        if (!RANGE.test(target)) return ok(action, ctx.excel?.tables.length ? "bad-table" : "bad-address");
        return ok(action, rangeBox(target).r1 === rangeBox(target).r0 ? "invalid" : undefined);
      }
      case "conditionalFormat":
        if (!RANGE.test(action.address)) return ok(action, "bad-address");
        return ok(action, (action.rule === "greaterThan" || action.rule === "lessThan") && action.value === undefined ? "invalid" : undefined);
      case "freezeRows":
        return ok(action);
      case "addPivotTable": {
        if (!CELL.test(action.destination.split("!").pop()!.replace(/\$/g, ""))) return ok(action, "bad-address");
        const table = (ctx.excel?.tables ?? []).find((t) => t.name === action.source);
        if (!table && !RANGE.test(action.source)) return ok(action, ctx.excel?.tables.length ? "bad-table" : "bad-address");
        if (!table && rangeBox(action.source).r1 === rangeBox(action.source).r0) return ok(action, "invalid");
        // Felterne skal være overskrifter i kilden. De kendes for tabeller og for et område, der
        // begynder i arkets første brugte række; ellers afgør Excel det ved udførelsen.
        const headers = table?.headers ?? headersOf(ctx, action.source);
        const fields = [...action.rows, ...action.columns, ...action.values.map((v) => v.field)];
        if (headers && fields.some((f) => !headers.some((h) => h.trim().toLowerCase() === f.trim().toLowerCase()))) return ok(action, "bad-field");
        // Samme felt kan ikke både være række og kolonne.
        return ok(action, action.rows.some((r) => action.columns.includes(r)) ? "invalid" : undefined);
      }
      case "setValidation": {
        if (!RANGE.test(action.address)) return ok(action, "bad-address");
        if (cellCount(action.address) > MAX_AGENT_CELLS * 10) return ok(action, "too-large");
        const numeric = action.min !== undefined || action.max !== undefined;
        // Enten en liste eller en talgrænse; listens værdier adskilles med komma i Excel og må ikke selv indeholde et.
        if (!!action.list === numeric) return ok(action, "invalid");
        if (action.list?.some((v) => v.includes(","))) return ok(action, "invalid");
        return ok(action, action.min !== undefined && action.max !== undefined && action.min > action.max ? "invalid" : undefined);
      }
      case "saveSheet": {
        if (!opts.export) return ok(action, "no-bridge");
        return ok(action, EXPORT_FILE_NAME.test(action.fileName) ? undefined : "invalid");
      }
      case "replace":
      case "comment":
        return ok(action, inDoc(action.find) ? undefined : "not-found");
      case "delete":
        if ((action.paragraph === undefined) === (action.find === undefined)) return ok(action, "invalid");
        return ok(action, action.find !== undefined ? (inDoc(action.find) ? undefined : "not-found") : para(action.paragraph, true));
      case "move":
        if (action.paragraph === action.to) return ok(action, "invalid");
        return ok(action, para(action.paragraph, true) ?? para(action.to, true));
      case "setStyle":
        return ok(action, para(action.paragraph, true));
      case "insert":
        if (action.position === "before" || action.position === "after") return ok(action, para(action.paragraph, true));
        return ok(action);
      case "insertTable": {
        // Alle rækker skal have lige mange celler; ellers er det uklart, hvor cellerne hører til.
        if (action.rows.some((r) => r.length !== action.rows[0].length)) return ok(action, "invalid");
        if (action.position === "before" || action.position === "after") return ok(action, para(action.paragraph, true));
        return ok(action);
      }
      case "addTableRow":
      case "deleteTableRow": {
        if (!paras) return ok(action, "no-structure");
        const p = paras[action.paragraph - 1];
        if (!p) return ok(action, "bad-paragraph");
        if (!p.table && !p.cell) return ok(action, "not-in-table");
        if (action.action === "deleteTableRow") {
          // Tabellens eneste række kan ikke slettes som række (så skal hele tabellen væk).
          const rows = p.cell ? new Set(paras.filter((x) => x.cell?.table === p.cell!.table).map((x) => x.cell!.row)).size : 2;
          return ok(action, rows < 2 ? "invalid" : undefined);
        }
        // Flere værdier end tabellen har kolonner, kan ikke placeres.
        const cols = p.cell ? Math.max(...paras.filter((x) => x.cell?.table === p.cell!.table).map((x) => x.cell!.col)) : undefined;
        return ok(action, cols !== undefined && action.values.length > cols ? "invalid" : undefined);
      }
      case "setText": {
        if (!slides) return ok(action, "no-structure");
        const [sn, shn] = action.shape.split(".").map(Number);
        const sh = slides[sn - 1]?.shapes[shn - 1];
        return ok(action, !slides[sn - 1] ? "bad-slide" : !sh ? "bad-shape" : sh.kind === "other" ? "invalid" : undefined);
      }
      case "addSlide":
        if (!slides) return ok(action, "no-structure");
        if (action.after > slides.length) return ok(action, "bad-slide");
        return ok(action, action.layout && ctx.layouts && !ctx.layouts.includes(action.layout) ? "bad-layout" : undefined);
      case "deleteSlide":
        return ok(action, !slides ? "no-structure" : slideOk(action.slide) ? undefined : "bad-slide");
      case "moveSlide":
        if (!slides) return ok(action, "no-structure");
        if (action.slide === action.to) return ok(action, "invalid");
        return ok(action, slideOk(action.slide) && slideOk(action.to) ? undefined : "bad-slide");
      case "appendRows": {
        const t = ctx.excel?.tables.find((x) => x.name === action.table);
        if (!t) return ok(action, "bad-table");
        if (action.values.some((r) => r.length > t.headers.length)) return ok(action, "invalid");
        return ok(action, action.values.reduce((n, r) => n + r.length, 0) > MAX_AGENT_CELLS ? "too-large" : undefined);
      }
      case "setFormula":
        return CELL.test(action.address.split("!").pop()!) ? { action } : { action, problem: "bad-address" };
      case "setValues": {
        if (!RANGE.test(action.address)) return { action, problem: "bad-address" };
        const cells = action.values.reduce((n, r) => n + r.length, 0);
        return cells > MAX_AGENT_CELLS ? { action, problem: "too-large" } : { action };
      }
      case "clear":
      case "setNumberFormat":
        if (!RANGE.test(action.address)) return { action, problem: "bad-address" };
        return cellCount(action.address) > MAX_AGENT_CELLS ? { action, problem: "too-large" } : { action };
      case "autofitColumns":
        return ok(action, RANGE.test(action.address) ? undefined : "bad-address");
      case "deleteRows": {
        if (!RANGE.test(action.address)) return ok(action, "bad-address");
        const b = rangeBox(action.address);
        if (b.r1 - b.r0 + 1 > MAX_AGENT_ROWS) return ok(action, "too-large");
        // Kun hele tabelrækker i en tabels data: overskriften eller rækker, der stikker ud af tabellen, ville ødelægge den.
        return ok(action, tablesOn(ctx, action.address).some((t) => overlaps(b, t.box) && !(b.r0 > t.box.r0 && b.r1 <= t.box.r1)) ? "in-table" : undefined);
      }
      case "deleteEmptyRows":
      case "formatTable": {
        if (tablesOf(ctx).some((t) => t.name === action.target)) return ok(action);
        if (!RANGE.test(action.target)) return ok(action, ctx.excel?.tables.length ? "bad-table" : "bad-address");
        if (action.action === "deleteEmptyRows") return ok(action);
        // En ny tabel må ikke overlappe en eksisterende (Excel afviser det).
        const b = rangeBox(action.target);
        return ok(action, b.r1 === b.r0 ? "invalid" : tablesOn(ctx, action.target).some((t) => overlaps(b, t.box)) ? "in-table" : undefined);
      }
    }
  });
}
