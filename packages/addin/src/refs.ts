import { rangeBox, type AgentAction, type DocContext, type HostKind } from "@miai/shared";
import { EXCEL_ADDRESS } from "./excelAddress";

/**
 * Henvisninger til steder i dokumentet (#22): celler og områder i Excel, slides i PowerPoint og
 * – i agentens plan – afsnit og tekst i Word. De gøres klikbare, så brugeren kan hoppe til stedet.
 */
export interface DocRef {
  /** Excel: celle eller område, evt. med arknavn ("Salg!C2:C5"). */
  address?: string;
  /** Excel: en tabel ved navn. */
  table?: string;
  /** PowerPoint: slidenummer (1 = første). */
  slide?: number;
  /** Word: afsnitsnummer i hele dokumentet (1 = første). */
  paragraph?: number;
  /** Word: tekst, der står i dokumentet. */
  find?: string;
}

export interface RefMatch {
  start: number;
  end: number;
  ref: DocRef;
}

const SHEET = "(?:'[^'\\n]{1,31}'|[\\p{L}\\p{N}_.]{1,31})!";
const CELL = "\\$?[A-Z]{1,3}\\$?[1-9][0-9]{0,6}";
/** Celle eller område, evt. med arknavn; ikke midt i et ord eller tal og ikke foran en bindestreg og et ord ("A4-papir"). */
const EXCEL_RE = new RegExp(`(?<![\\p{L}\\p{N}_$!'.])(${SHEET})?(${CELL})(:${CELL})?(?![\\p{L}\\p{N}_(!]|-\\p{L})`, "gu");
/** "slide 3" / "dias 3" (dansk og engelsk). */
const SLIDE_RE = /(?<![\p{L}\p{N}])(?:slides?|dias(?:set)?)\s+(\d{1,3})(?![\p{L}\p{N}])/giu;

/**
 * Finder henvisninger i en tekst. I Excel linkes områder (A1:C5) og adresser med arknavn altid; en
 * enkelt celle (B7) kun, når den ligger inden for arkets brugte område – ellers ville "Q1 2026"
 * eller "A4-papir" blive til links. `used` er arkets brugte område uden arknavn (fx "A1:E30").
 */
export function findRefs(text: string, host: HostKind, opts: { used?: string } = {}): RefMatch[] {
  const out: RefMatch[] = [];
  if (host === "excel") {
    const used = opts.used ? rangeBox(opts.used) : undefined;
    for (const m of text.matchAll(EXCEL_RE)) {
      const [whole, sheet, from, to] = m;
      if (!sheet && !to) {
        if (!used) continue;
        const b = rangeBox(from);
        if (b.r0 < used.r0 || b.r0 > used.r1 || b.c0 < used.c0 || b.c0 > used.c1) continue;
      }
      out.push({ start: m.index, end: m.index + whole.length, ref: { address: whole.replace(/\$/g, "") } });
    }
  } else if (host === "powerpoint") {
    for (const m of text.matchAll(SLIDE_RE)) out.push({ start: m.index, end: m.index + m[0].length, ref: { slide: Number(m[1]) } });
  }
  return out;
}

/** Arkets brugte område fra oversigten i konteksten ("… used range A1:E30 …"). */
export function usedRangeOf(ctx: Pick<DocContext, "surroundings" | "excel"> | null | undefined): string | undefined {
  const active = ctx?.excel?.sheets?.find((s) => s.name === ctx.excel?.activeSheet)?.used;
  return active ?? /used range ([A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?)/.exec(ctx?.surroundings ?? "")?.[1];
}

/**
 * Gør henvisninger i renderet Markdown klikbare: tekstnoder uden for kode, links og figurer deles
 * op, og hver henvisning bliver en knap med data-ref. Teksten ændres ikke.
 */
export function linkifyRefs(root: HTMLElement, find: (text: string) => RefMatch[]): void {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.parentElement?.closest("code, pre, a, button, .figure") ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  const nodes: Text[] = [];
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  for (const node of nodes) {
    const text = node.data;
    const matches = find(text);
    if (!matches.length) continue;
    const frag = document.createDocumentFragment();
    let at = 0;
    for (const m of matches) {
      if (m.start < at) continue;
      frag.append(text.slice(at, m.start));
      const b = document.createElement("button");
      b.type = "button";
      b.className = "docref";
      b.dataset.ref = JSON.stringify(m.ref);
      b.textContent = text.slice(m.start, m.end);
      frag.append(b);
      at = m.end;
    }
    frag.append(text.slice(at));
    node.replaceWith(frag);
  }
}

/** Stedet i dokumentet, en handling i agentens plan gælder – eller undefined, hvis det ikke kan vises. */
export function targetOf(a: AgentAction, ctx: DocContext): DocRef | undefined {
  if (ctx.host === "excel") {
    const ref =
      a.action === "addPivotTable"
        ? a.source
        : "address" in a
          ? a.address
          : "target" in a
            ? a.target
            : "source" in a
              ? a.source
              : "table" in a
                ? a.table
                : undefined;
    if (!ref) return undefined;
    return EXCEL_ADDRESS.test(ref) ? { address: ref.replace(/\$/g, "") } : { table: ref };
  }
  if (ctx.host === "powerpoint") {
    if (a.action === "deleteSlide" || a.action === "moveSlide") return { slide: a.slide };
    if (a.action === "setText") return { slide: Number(a.shape.split(".")[0]) };
    if (a.action === "addSlide") return a.after >= 1 ? { slide: a.after } : undefined;
    if (a.action === "replace") {
      const i = ctx.slides?.findIndex((sl) => sl.shapes.some((sh) => sh.text.includes(a.find))) ?? -1;
      return i >= 0 ? { slide: i + 1 } : undefined;
    }
    return undefined;
  }
  if (ctx.host === "word") {
    // Afsnitsnumre gælder kun hele dokumentet; for en markering kan stedet ikke findes igen med sikkerhed.
    const para = (n: number | undefined) => (n !== undefined && ctx.isWholeDocument && ctx.paragraphs?.[n - 1] ? { paragraph: n } : undefined);
    if (a.action === "replace" || a.action === "comment") return { find: a.find };
    if (a.action === "delete") return a.find ? { find: a.find } : para(a.paragraph);
    if (a.action === "setStyle" || a.action === "move" || a.action === "insert" || a.action === "insertTable") return para(a.paragraph);
    if (a.action === "addTableRow" || a.action === "deleteTableRow") return para(a.paragraph);
  }
  return undefined;
}
