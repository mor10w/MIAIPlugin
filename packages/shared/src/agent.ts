import { z } from "zod";
import type { DocContext, HostKind } from "./types.js";

/**
 * Agent-tilstand (#6, ADR-0012): modellen foreslår en plan af handlinger fra en fast whitelist.
 * Intet udføres af motoren. Add-in'et viser planen, brugeren vælger og bekræfter, og først
 * derefter udføres handlingerne med Office.js – med fortryd for hele sessionen.
 *
 * Bevidst ingen netværks-, fil- eller formateringsværktøjer: dokumentindhold kan indeholde
 * prompt injection, og det værste en kapret plan kan gøre, er at foreslå tekstændringer,
 * som brugeren ser, før de udføres.
 */

/** Maks. antal handlinger i én plan (loft pr. session). */
export const MAX_AGENT_ACTIONS = 20;
/** Maks. antal celler, én Excel-handling må skrive. */
export const MAX_AGENT_CELLS = 2_000;

const reason = z.string().max(200).describe("Very short reason shown to the user, in the user's language");
const find = z.string().min(1).max(255).describe("Exact text copied verbatim from the document (a few words, unique if possible)");

export const WordStyle = z.enum(["normal", "heading1", "heading2", "heading3", "bullet"]);
export type WordStyle = z.infer<typeof WordStyle>;

export const AgentAction = z.discriminatedUnion("action", [
  // --- Word ------------------------------------------------------------------
  z.object({
    action: z.literal("replace"),
    find,
    replace: z.string().max(4_000),
    all: z.boolean().optional().describe("Replace every occurrence (default: only the first)"),
    reason,
  }),
  z.object({
    action: z.literal("insert"),
    text: z.string().min(1).max(4_000).describe("New paragraph text (plain text, no markdown)"),
    position: z.enum(["start", "end", "after"]).describe("start/end of the document, or after the paragraph containing `anchor`"),
    anchor: z.string().max(255).optional().describe("Required when position is 'after': exact text from the document"),
    style: WordStyle.optional(),
    reason,
  }),
  z.object({ action: z.literal("delete"), find, reason }),
  z.object({
    action: z.literal("move"),
    find,
    anchor: z.string().min(1).max(255).describe("Exact text from the document to move next to"),
    position: z.enum(["before", "after"]),
    reason,
  }),
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

export const WORD_ACTIONS: AgentActionKind[] = ["replace", "insert", "delete", "move", "comment"];
export const EXCEL_ACTIONS: AgentActionKind[] = ["setValues", "setFormula", "clear"];
export const agentActionsFor = (host: HostKind): AgentActionKind[] => (host === "word" ? WORD_ACTIONS : host === "excel" ? EXCEL_ACTIONS : []);

/** Destruktive handlinger fremhæves i planen. */
export const isDestructive = (a: AgentAction) => a.action === "delete" || a.action === "clear" || (a.action === "replace" && !a.replace);

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
    "Available actions (Word):",
    "- replace {find, replace, all?}: replace text. `find` must be copied exactly from the document.",
    "- insert {text, position: start|end|after, anchor?, style?}: insert a new paragraph. style is one of normal, heading1, heading2, heading3, bullet.",
    "- delete {find}: delete text copied exactly from the document.",
    "- move {find, anchor, position: before|after}: move text next to the anchor text.",
    "- comment {find, text}: add a review comment on the text.",
    "`find` and `anchor` must be short (a few words up to one sentence) and copied verbatim, so they can be found with search.",
  ].join("\n");
}

export interface CheckedAction {
  action: AgentAction;
  /** Udfyldt, hvis handlingen ikke kan udføres (vises i planen, og handlingen fravælges). */
  problem?: "wrong-host" | "not-found" | "bad-address" | "too-large" | "anchor-missing";
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
export function checkPlan(plan: AgentPlan, ctx: Pick<DocContext, "host" | "text">): CheckedAction[] {
  const allowed = agentActionsFor(ctx.host);
  const text = ctx.text ?? "";
  const inDoc = (s?: string) => !!s && text.includes(s);
  return plan.actions.slice(0, MAX_AGENT_ACTIONS).map((action): CheckedAction => {
    if (!allowed.includes(action.action)) return { action, problem: "wrong-host" };
    switch (action.action) {
      case "replace":
      case "delete":
      case "comment":
        return inDoc(action.find) ? { action } : { action, problem: "not-found" };
      case "move":
        return !inDoc(action.find) ? { action, problem: "not-found" } : !inDoc(action.anchor) ? { action, problem: "anchor-missing" } : { action };
      case "insert":
        return action.position === "after" && !inDoc(action.anchor) ? { action, problem: "anchor-missing" } : { action };
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
