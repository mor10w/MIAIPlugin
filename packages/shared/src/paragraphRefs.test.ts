import { describe, expect, it } from "vitest";
import { citableParagraphs, paragraphMark, renderCitableParagraphs } from "./paragraphRefs.js";
import { getTask } from "./tasks/catalog.js";
import { TaskRequest } from "./types.js";

/** Henvisninger til afsnit i Word-chatten (#77): dokumentet nummereres for modellen, og den bedes henvise med markøren. */

const paragraphs = [
  { text: "Tilbud", style: "Heading1" },
  { text: "", style: "Normal" },
  { text: "Prisen er 499 kr. pr. måned.", style: "Normal" },
];
const text = paragraphs.map((p) => p.text).join("\n");
const system = (context: Record<string, unknown>) =>
  getTask("chat")!.buildMessages(TaskRequest.parse({ context, params: { prompt: "Hvad koster det?" } }))[0].content as string;

const RULE = "cite the paragraph by its marker, e.g. [¶12]";

describe("afsnit med markør til chatten", () => {
  it("ét afsnit pr. linje med en kort markør; tomme afsnit udelades, og numrene følger dokumentet", () => {
    expect(paragraphMark(12)).toBe("[¶12]");
    expect(renderCitableParagraphs(paragraphs)).toBe("[¶1] Tilbud\n[¶3] Prisen er 499 kr. pr. måned.");
  });

  it("en tabelrække står på én linje med markøren for rækkens første afsnit", () => {
    const cell = (table: number, row: number, col: number, t: string) => ({ text: t, style: "Normal", table: true, cell: { table, row, col } });
    const doc = [
      { text: "Priser", style: "Heading2" },
      cell(1, 1, 1, "Plan"),
      cell(1, 1, 2, "Pris"),
      cell(1, 2, 1, ""),
      cell(1, 2, 2, "799 kr"),
      cell(1, 2, 2, "ekskl. moms"),
      { text: "Alle priser er pr. måned.", style: "Normal" },
    ];
    expect(renderCitableParagraphs(doc)).toBe("[¶1] Priser\n[¶2] Plan | Pris\n[¶4]  | 799 kr ekskl. moms\n[¶7] Alle priser er pr. måned.");
  });

  it("kun når hele Word-dokumentet er konteksten", () => {
    expect(citableParagraphs({ host: "word", isWholeDocument: true, paragraphs })).toBe(paragraphs);
    expect(citableParagraphs({ host: "word", isWholeDocument: false, paragraphs })).toBeUndefined();
    expect(citableParagraphs({ host: "word", isWholeDocument: true })).toBeUndefined();
    expect(citableParagraphs({ host: "word", isWholeDocument: true, paragraphs: [{ text: " ", style: "Normal" }] })).toBeUndefined();
    expect(citableParagraphs({ host: "powerpoint", isWholeDocument: true, paragraphs })).toBeUndefined();
  });
});

describe("chattens prompt", () => {
  it("Word med hele dokumentet: nummererede afsnit og reglen om at henvise med markøren", () => {
    const sys = system({ host: "word", text, isWholeDocument: true, paragraphs });
    expect(sys).toContain("<document>\n[¶1] Tilbud\n[¶3] Prisen er 499 kr. pr. måned.\n</document>");
    expect(sys).toContain(RULE);
    expect(sys).toContain("the user can click it to jump to the paragraph");
    // Dokumentet står der én gang: de nummererede afsnit erstatter teksten.
    expect(sys.match(/Prisen er 499 kr/g)).toHaveLength(1);
  });

  it("et afkortet dokument er stadig markeret som afkortet", () => {
    const sys = system({ host: "word", text: `${text}\n[… afkortet]`, isWholeDocument: true, truncated: true, originalChars: 90_000, paragraphs });
    expect(sys).toContain("[¶3] Prisen er 499 kr. pr. måned.\n[… afkortet]\n</document>");
  });

  it("Word uden afsnit eller med en markering: teksten som før, ingen regel om markører", () => {
    for (const context of [
      { host: "word", text, isWholeDocument: true },
      { host: "word", text, isWholeDocument: false, paragraphs },
    ]) {
      const sys = system(context);
      expect(sys).toContain(`The user's current ${context.isWholeDocument ? "document" : "selection"}:\n<document>\n${text}\n</document>`);
      expect(sys).not.toContain("¶");
    }
  });

  it("Excel og PowerPoint er uændrede", () => {
    const excel = system({ host: "excel", text: "Region\tOmsætning\nNord\t120", address: "Ark1!A1:B2" });
    expect(excel).toContain("Refer to cells by their addresses (B7, C2:C5, or Budget!B2 on another sheet) – the user can click an address to jump to it.");
    expect(excel).toContain("The user's current selection (Ark1!A1:B2):\n<document>\nRegion\tOmsætning\nNord\t120\n</document>");
    const ppt = system({ host: "powerpoint", text: "Titel\nPunkt", isWholeDocument: true });
    expect(ppt).toContain("Refer to slides by number ('slide 3') – the user can click it to go to the slide.");
    expect(ppt).toContain("The user's current document:\n<document>\nTitel\nPunkt\n</document>");
    for (const sys of [excel, ppt]) expect(sys).not.toContain("¶");
  });
});
