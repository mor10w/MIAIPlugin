import { z } from "zod";
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
 * de indbyggede tabeltypografier, talformater og kolonnebredde.
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
    action: z.literal("autofitColumns"),
    address: z.string().min(2).max(40).describe("Range whose columns get a width that fits their content, e.g. 'A1:F40'"),
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

export const WORD_ACTIONS: AgentActionKind[] = ["replace", "insert", "delete", "move", "setStyle", "comment"];
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
];
export const POWERPOINT_ACTIONS: AgentActionKind[] = ["replace", "setText", "addSlide", "deleteSlide", "moveSlide"];
export const agentActionsFor = (host: HostKind): AgentActionKind[] =>
  host === "word" ? WORD_ACTIONS : host === "excel" ? EXCEL_ACTIONS : host === "powerpoint" ? POWERPOINT_ACTIONS : [];

/** Destruktive handlinger fremhæves i planen. */
export const isDestructive = (a: AgentAction) =>
  a.action === "delete" ||
  a.action === "clear" ||
  a.action === "deleteSlide" ||
  a.action === "deleteRows" ||
  a.action === "deleteEmptyRows" ||
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
  if (x.tables.length) lines.push("Tables:", ...x.tables.map((t) => `- ${t.name} (${t.address}): ${t.headers.join(" | ")}`));
  if (x.columnFormats.length) lines.push(`Number formats: ${x.columnFormats.join("; ")}`);
  return lines.join("\n");
}

/**
 * Dokumentets afsnit som tekst til prompten: "¶3 [Heading2] Priser", med (list) og (table).
 * Lange afsnit vises i fuld længde, så modellen kan citere ordret.
 */
export function renderParagraphs(paras: NonNullable<DocContext["paragraphs"]>): string {
  return paras.map((p, i) => `¶${i + 1} [${p.style}]${p.list ? " (list)" : ""}${p.table ? " (table)" : ""} ${p.text}`).join("\n");
}

/** Beskrivelse af værktøjerne til systemprompten – kun dem, der gælder for værten. */
export function agentToolGuide(host: HostKind): string {
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
      "For 'make the table nice/modern/readable': formatTable, then setNumberFormat for number and date columns where needed, then autofitColumns. Do not change cell values for this.",
      "Formatting is preserved: values never change cell formats, and new cells directly below or next to data get the neighbouring cells' formats. Fonts, colours and borders come only from the table style – there is no action for free formatting.",
      "Use addresses that match the sheet overview; row numbers are shown in the first column of the selected values. Never write over existing data unless the user asked for it. Addresses refer to the sheet before any of your actions – the add-in deletes rows last, from the bottom up.",
    ].join("\n");
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
    "Available actions (Word). The document is shown as numbered paragraphs: ¶number [style] text; (list) = list item, (table) = inside a table.",
    "- replace {find, replace, all?}: change text inside a paragraph. `find` is copied verbatim from one paragraph; only the words that differ are replaced, so bold, links and other formatting around them are kept.",
    "- insert {text, position: before|after|start|end, paragraph?, style?}: add new paragraph(s) before/after ¶paragraph or at the start/end. One line per paragraph; use Markdown for structure and emphasis ('## ' heading, '- ' bullet, '1. ' numbered, **bold**, *italic*, [text](https://…)). Lines without a prefix get `style` (default normal) – the new text never inherits the neighbour's formatting.",
    "- delete {paragraph} removes a whole paragraph; delete {find} removes only that text inside a paragraph.",
    "- move {paragraph, to, position: before|after}: move a whole paragraph with its formatting.",
    "- setStyle {paragraph, style}: change a paragraph's role (e.g. body text → heading2, or a line → bullet) without touching its text or inline formatting.",
    "- comment {find, text}: add a review comment on the text.",
    "Styles: normal, heading1, heading2, heading3, bullet, numbered, quote. Use the document's existing structure: match the heading levels and list style already used around the place you edit.",
    "Preserve formatting: prefer small replace actions over rewriting whole paragraphs; use setStyle instead of delete + insert to change a role; never delete, move or insert inside paragraphs marked (table) – edit their text with replace only.",
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
  /** Handlingen er ufuldstændig (fx delete uden afsnit og tekst). */
  | "invalid"
  /** Word uden afsnitsstruktur (før WordApi 1.3): afsnitshandlinger kan ikke bruges. */
  | "no-structure"
  /** PowerPoint: slide eller figur findes ikke; layout findes ikke i præsentationen. */
  | "bad-slide"
  | "bad-shape"
  | "bad-layout"
  /** Excel: tabellen findes ikke på arket. */
  | "bad-table";

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
const tablesOf = (ctx: Pick<DocContext, "excel">) => (ctx.excel?.tables ?? []).map((t) => ({ name: t.name, box: rangeBox(t.address) }));

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
export function checkPlan(plan: AgentPlan, ctx: Pick<DocContext, "host" | "text" | "paragraphs" | "slides" | "layouts" | "excel">): CheckedAction[] {
  const allowed = agentActionsFor(ctx.host);
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
    switch (action.action) {
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
        return ok(action, tablesOf(ctx).some((t) => overlaps(b, t.box) && !(b.r0 > t.box.r0 && b.r1 <= t.box.r1)) ? "in-table" : undefined);
      }
      case "deleteEmptyRows":
      case "formatTable": {
        if (tablesOf(ctx).some((t) => t.name === action.target)) return ok(action);
        if (!RANGE.test(action.target)) return ok(action, ctx.excel?.tables.length ? "bad-table" : "bad-address");
        if (action.action === "deleteEmptyRows") return ok(action);
        // En ny tabel må ikke overlappe en eksisterende (Excel afviser det).
        const b = rangeBox(action.target);
        return ok(action, b.r1 === b.r0 ? "invalid" : tablesOf(ctx).some((t) => overlaps(b, t.box)) ? "in-table" : undefined);
      }
    }
  });
}
