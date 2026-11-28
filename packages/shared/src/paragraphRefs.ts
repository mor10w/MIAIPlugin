import type { DocContext } from "./types.js";

type Paragraphs = NonNullable<DocContext["paragraphs"]>;

/**
 * Henvisninger til afsnit i Word-chatten (#77): dokumentet vises for modellen med en kort markør
 * foran hvert afsnit, og modellen henviser med den samme markør. Add-in'et gør markøren klikbar.
 */

/** Markøren for afsnit n, fx "[¶12]". Kort, fordi den står foran hvert afsnit og koster tokens. */
export const paragraphMark = (n: number): string => `[¶${n}]`;

/**
 * Afsnittene, et chatsvar kan henvise til – eller undefined. Kun når hele Word-dokumentet er
 * konteksten: numrene gælder hele dokumentet, og for en markering ville Word give de hele afsnit,
 * også den tekst, brugeren ikke har markeret.
 */
export function citableParagraphs(ctx: Pick<DocContext, "host" | "isWholeDocument" | "paragraphs">): Paragraphs | undefined {
  return ctx.host === "word" && ctx.isWholeDocument && ctx.paragraphs?.some((p) => p.text.trim()) ? ctx.paragraphs : undefined;
}

/**
 * Dokumentet til chatten: "[¶3] Priser", ét afsnit pr. linje. Tomme afsnit udelades (numrene følger
 * stadig dokumentet), og en tabelrække står på én linje med markøren for rækkens første afsnit og
 * " | " mellem cellerne – så koster en tabel ikke en markør pr. celle.
 */
export function renderCitableParagraphs(paras: Paragraphs): string {
  const lines: string[] = [];
  let prev: Paragraphs[number]["cell"];
  paras.forEach((p, i) => {
    const c = p.cell;
    if (c && prev && c.table === prev.table && c.row === prev.row) lines[lines.length - 1] += `${c.col === prev.col ? " " : " | "}${p.text}`;
    else if (c || p.text.trim()) lines.push(`${paragraphMark(i + 1)} ${p.text}`);
    prev = c;
  });
  return lines.join("\n");
}
