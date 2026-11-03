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
 * dokumentets egne typografier og begrænset Markdown (fed, kursiv, links, lister).
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
  // --- Excel -----------------------------------------------------------------
  z.object({
    action: z.literal("setValues"),
    address: z.string().min(2).max(40).describe("Top-left cell or range on the active sheet, e.g. 'D2' or 'D2:E5'"),
    values: z
      .array(z.array(z.string().max(1_000)))
      .min(1)
      .max(200)
      .describe("Rows of cell values written from the top-left cell"),
    reason,
  }),
  z.object({
    action: z.literal("setFormula"),
    address: z.string().min(2).max(20).describe("One cell, e.g. 'F2'"),
    formula: z.string().min(2).max(1_000).describe("Excel formula starting with '=', English function names, comma separator"),
    reason,
  }),
  z.object({ action: z.literal("clear"), address: z.string().min(2).max(40).describe("Range on the active sheet whose contents are cleared"), reason }),
]);
export type AgentAction = z.infer<typeof AgentAction>;
export type AgentActionKind = AgentAction["action"];

export const AgentPlan = z.object({
  summary: z.string().max(1_000).describe("One or two sentences: what the plan does, or why it cannot be done with the available actions"),
  actions: z.array(AgentAction).max(MAX_AGENT_ACTIONS),
});
export type AgentPlan = z.infer<typeof AgentPlan>;

export const WORD_ACTIONS: AgentActionKind[] = ["replace", "insert", "delete", "move", "setStyle", "comment"];
export const EXCEL_ACTIONS: AgentActionKind[] = ["setValues", "setFormula", "clear"];
export const agentActionsFor = (host: HostKind): AgentActionKind[] => (host === "word" ? WORD_ACTIONS : host === "excel" ? EXCEL_ACTIONS : []);

/** Destruktive handlinger fremhæves i planen. */
export const isDestructive = (a: AgentAction) => a.action === "delete" || a.action === "clear" || (a.action === "replace" && !a.replace);

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
      '- setValues {address, values}: write a 2D array of values starting at address (e.g. "D2"). Use it for new columns, labels and data.',
      '- setFormula {address, formula}: put ONE formula in ONE cell (e.g. "F2"). For a column of formulas, use one action per cell (max 20 actions).',
      "- clear {address}: clear the contents of a range. Only when the user asks to remove data.",
      "Use addresses that match the sheet overview. Never write over existing data unless the user asked for it.",
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
  | "no-structure";

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

/**
 * Tjekker planen mod værten og dokumentet, før den vises: handlinger til den forkerte vært,
 * tekst, der ikke findes i dokumentet, ugyldige adresser og for store skrivninger markeres.
 */
export function checkPlan(plan: AgentPlan, ctx: Pick<DocContext, "host" | "text" | "paragraphs">): CheckedAction[] {
  const allowed = agentActionsFor(ctx.host);
  const paras = ctx.paragraphs;
  const text = paras ? paras.map((p) => p.text).join("\n") : (ctx.text ?? "");
  // Søgning i Word sker inden for ét afsnit ad gangen.
  const inDoc = (s?: string) => !!s && !s.includes("\n") && (paras ? paras.some((p) => p.text.includes(s)) : text.includes(s));
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
      case "setFormula":
        return CELL.test(action.address.split("!").pop()!) ? { action } : { action, problem: "bad-address" };
      case "setValues": {
        if (!RANGE.test(action.address)) return { action, problem: "bad-address" };
        const cells = action.values.reduce((n, r) => n + r.length, 0);
        return cells > MAX_AGENT_CELLS ? { action, problem: "too-large" } : { action };
      }
      case "clear":
        if (!RANGE.test(action.address)) return { action, problem: "bad-address" };
        return cellCount(action.address) > MAX_AGENT_CELLS ? { action, problem: "too-large" } : { action };
    }
  });
}
