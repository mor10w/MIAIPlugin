/**
 * Tekst-hjælpere til agentens Word-handlinger (#6, ADR-0012): begrænset Markdown til ny tekst
 * og den mindst mulige ændring ved erstat, så formatering bevares, hvor intet ændres.
 */

export type BlockStyle = "normal" | "heading1" | "heading2" | "heading3" | "bullet" | "numbered" | "quote";

export interface Span {
  start: number;
  end: number;
  bold?: boolean;
  italic?: boolean;
  /** Kun http(s) og mailto – andre skemaer (javascript:, file:) ignoreres. */
  link?: string;
}

export interface InlineText {
  /** Teksten uden Markdown-tegn. */
  text: string;
  /** Formaterede stykker (forskydninger i `text`). */
  spans: Span[];
}

const SAFE_LINK = /^(https?:\/\/|mailto:)[^\s]+$/i;

/**
 * Inline-Markdown: **fed**, *kursiv* / _kursiv_, ***begge*** og [tekst](url). Alt andet står
 * som det er. Ingen HTML – teksten indsættes som tekst, og formateringen sættes bagefter.
 */
export function parseInline(md: string): InlineText {
  const spans: Span[] = [];
  let out = "";
  const re =
    /\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|(?<![\p{L}\p{N}*])\*(?!\s)(.+?)(?<!\s)\*(?![\p{L}\p{N}*])|(?<![\p{L}\p{N}_])_(?!\s)(.+?)(?<!\s)_(?![\p{L}\p{N}_])|\[([^\]]+)\]\(([^)\s]+)\)/gu;
  let last = 0;
  for (const m of md.matchAll(re)) {
    out += md.slice(last, m.index);
    const start = out.length;
    if (m[1] !== undefined) {
      out += m[1];
      spans.push({ start, end: out.length, bold: true, italic: true });
    } else if (m[2] !== undefined) {
      out += m[2];
      spans.push({ start, end: out.length, bold: true });
    } else if (m[3] !== undefined || m[4] !== undefined) {
      out += m[3] ?? m[4];
      spans.push({ start, end: out.length, italic: true });
    } else {
      out += m[5];
      if (SAFE_LINK.test(m[6])) spans.push({ start, end: out.length, link: m[6] });
    }
    last = m.index! + m[0].length;
  }
  out += md.slice(last);
  return { text: out, spans };
}

/**
 * Blok-Markdown: én linje = ét afsnit. "# ", "## ", "### " = overskrifter, "- " / "* " = punkt,
 * "1. " = nummereret, "> " = citat. Linjer uden præfiks får `defaultStyle`. Tomme linjer springes over.
 */
export function parseBlocks(text: string, defaultStyle: BlockStyle = "normal"): { style: BlockStyle; inline: InlineText }[] {
  const out: { style: BlockStyle; inline: InlineText }[] = [];
  for (const raw of text.replace(/\r\n?/g, "\n").split("\n")) {
    const line = raw.trimEnd();
    if (!line.trim()) continue;
    let m: RegExpMatchArray | null;
    let style = defaultStyle;
    let body = line.trim();
    if ((m = body.match(/^(#{1,3})\s+(.*)$/))) {
      style = (["heading1", "heading2", "heading3"] as const)[m[1].length - 1];
      body = m[2];
    } else if ((m = body.match(/^[-*•]\s+(.*)$/))) {
      style = "bullet";
      body = m[1];
    } else if ((m = body.match(/^\d+[.)]\s+(.*)$/))) {
      style = "numbered";
      body = m[1];
    } else if ((m = body.match(/^>\s?(.*)$/))) {
      style = "quote";
      body = m[1];
    }
    out.push({ style, inline: parseInline(body) });
  }
  return out;
}

/**
 * Den mindst mulige ændring mellem `find` og `replace` på ordgrænser: fælles begyndelse og
 * slutning bevares (og dermed deres formatering), kun den midterste del erstattes.
 * "Pris: **499** kr pr. md." → "Pris: 599 kr pr. md." ændrer kun "499".
 *
 * Er den gamle del tom (ren indsættelse), udvides den med ét ord til venstre eller højre, så
 * der er noget at erstatte i dokumentet.
 */
export function minimalEdit(find: string, replace: string): { before: string; oldCore: string; newCore: string } {
  if (find === replace) return { before: find, oldCore: "", newCore: "" };
  const tokens = (s: string) => s.match(/\s+|[^\s]+/g) ?? [];
  const a = tokens(find);
  const b = tokens(replace);
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let q = 0;
  while (p + q < a.length && p + q < b.length && a[a.length - 1 - q] === b[b.length - 1 - q]) q++;
  let start = p;
  let endA = a.length - q;
  let endB = b.length - q;
  if (start === endA) {
    // Ren indsættelse: tag nabo-ordet med, så der er noget at erstatte i dokumentet.
    if (start > 0) {
      start--;
      if (/^\s+$/.test(a[start]) && start > 0) start--;
    } else if (endA < a.length) {
      endA++;
      endB++;
    }
  }
  return { before: a.slice(0, start).join(""), oldCore: a.slice(start, endA).join(""), newCore: b.slice(start, endB).join("") };
}

/** Antal gange `needle` forekommer i `hay` (uden overlap) – bruges til at vælge den rigtige forekomst. */
export function occurrences(hay: string, needle: string): number {
  if (!needle) return 0;
  let n = 0;
  for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + needle.length)) n++;
  return n;
}
