import { describe, expect, it } from "vitest";
import { AgentPlan, type DocContext } from "@miai/shared";
import { findRefs, locateParagraph, targetOf, usedRangeOf, withoutParagraphMarks } from "./refs";

const refs = (text: string, used?: string) => findRefs(text, "excel", { used }).map((m) => text.slice(m.start, m.end));

describe("findRefs (#22)", () => {
  it("Excel: områder og adresser med arknavn linkes altid", () => {
    expect(refs("Summen af C2:C5 er 760, og budgettet står i Budget!B2 og 'Mit ark'!A1:B3.")).toEqual(["C2:C5", "Budget!B2", "'Mit ark'!A1:B3"]);
    expect(findRefs("Se $C$2:$C$5", "excel")[0].ref).toEqual({ address: "C2:C5" });
  });

  it("Excel: arknavn med apostrof (to apostroffer) linkes som én adresse (#59)", () => {
    expect(refs("Tallene står i 'Bob''s'!A1:B3.")).toEqual(["'Bob''s'!A1:B3"]);
  });

  it("Excel: en enkelt celle linkes kun inden for arkets brugte område", () => {
    expect(refs("Cellen B7 er tom, men Q1 2026 og A4-papir er ikke celler.", "A1:E30")).toEqual(["B7"]);
    expect(refs("Cellen B7 er tom.")).toEqual([]);
    expect(refs("Z99 ligger uden for data.", "A1:E30")).toEqual([]);
  });

  it("Excel: ikke midt i ord, tal, funktioner eller filnavne", () => {
    expect(refs("SUM(C2:C5) og =LOG10(5), ABC123456789, v1.B2, CO2:N2O", "A1:E30")).toEqual(["C2:C5"]);
  });

  it("PowerPoint: slide og dias med nummer; andre værter har ingen henvisninger", () => {
    const t = "Se slide 3 og dias 12, men ikke slideshow 4.";
    expect(findRefs(t, "powerpoint").map((m) => m.ref)).toEqual([{ slide: 3 }, { slide: 12 }]);
    expect(findRefs("B2 og slide 3", "word")).toEqual([]);
  });

  it("usedRangeOf: fra arkoversigten eller agentens struktur", () => {
    expect(usedRangeOf({ surroundings: "Active sheet 'Ark1', used range A1:E30 (30 rows x 5 columns)" })).toBe("A1:E30");
    expect(usedRangeOf({ excel: { tables: [], columnFormats: [], activeSheet: "Salg", sheets: [{ name: "Salg", used: "B2:D9", headers: [] }] } })).toBe(
      "B2:D9",
    );
    expect(usedRangeOf(null)).toBeUndefined();
  });
});

describe("afsnit i Word (#77)", () => {
  const long = "Leverandøren skal senest fem arbejdsdage efter kontraktens underskrift fremsende en tidsplan for leverancen til kunden.";
  const paragraphs = [{ text: "Tilbud" }, { text: "" }, { text: `  ${long}` }, { text: "Pris" }];
  const found = (text: string, ps: { text: string }[] = paragraphs) =>
    findRefs(text, "word", { paragraphs: ps }).map((m) => [text.slice(m.start, m.end), m.ref]);

  it("markøren bliver til en henvisning med afsnittets nummer og begyndelsen af dets tekst", () => {
    expect(found("Fristen står i [¶3], og prisen i [¶4].")).toEqual([
      ["[¶3]", { paragraph: 3, excerpt: long.slice(0, 80) }],
      ["[¶4]", { paragraph: 4, excerpt: "Pris" }],
    ]);
  });

  it("også uden klammer og i en opremsning; et tomt afsnit har intet uddrag", () => {
    expect(found("Se [¶1, ¶4] og ¶ 2.").map(([text]) => text)).toEqual(["¶1", "¶4", "¶ 2"]);
    expect(found("Se [¶2].")).toEqual([["[¶2]", { paragraph: 2 }]]);
  });

  it("et nummer, dokumentet ikke har, bliver ikke til en henvisning", () => {
    expect(found("Se [¶5], [¶0], ¶12345 og [¶99].")).toEqual([]);
    // Uden afsnit fra dokumentet (fx et svar om en markering) linkes intet.
    expect(findRefs("Se [¶1].", "word")).toEqual([]);
    expect(found("Se [¶1].", [])).toEqual([]);
  });

  it("kilder ([1]) og tal er ikke afsnit", () => {
    expect(found("Ifølge [1] og § 3 er prisen 4 kr.")).toEqual([]);
  });

  describe("locateParagraph: afsnittet i dokumentet, som det er nu", () => {
    const doc = ["Tilbud", "Indledning", long, "Pris", "Pris", "Afslutning"];
    const excerpt = long.slice(0, 80);

    it("uændret dokument: nummeret gælder", () => {
      expect(locateParagraph(doc, 3, excerpt)).toBe(3);
      expect(locateParagraph(doc, 4, "Pris")).toBe(4);
    });

    it("et afsnit tidligere i dokumentet er slettet eller indsat: afsnittet findes på uddraget", () => {
      expect(locateParagraph(doc.slice(1), 3, excerpt)).toBe(2);
      expect(locateParagraph(["Nyt", "Nyt igen", ...doc], 3, excerpt)).toBe(5);
      // Uddraget er begyndelsen af afsnittet: resten må gerne være rettet, og blanktegn tæller ikke.
      expect(locateParagraph(["Indledning", `${excerpt.replace(" ", "\t  ")} – og mere til.`], 3, excerpt)).toBe(2);
    });

    it("flere afsnit med samme tekst: det nærmeste på det oprindelige nummer", () => {
      // "Pris" stod som afsnit 5; efter en sletning tidligere står de to som 3 og 4.
      expect(locateParagraph(["Tilbud", long, "Pris", "Pris", "Afslutning"], 5, "Pris")).toBe(4);
      // Et kort uddrag er hele afsnittet: "Priserne stiger" er et andet afsnit.
      expect(locateParagraph(["Priserne stiger", "Tilbud", "Pris"], 1, "Pris")).toBe(3);
    });

    it("uddraget findes ikke længere: stedet, hvor afsnittet stod – eller intet, hvis dokumentet er blevet kortere", () => {
      expect(locateParagraph(["Tilbud", "Indledning", "Helt ny tekst", "Pris"], 3, excerpt)).toBe(3);
      expect(locateParagraph(["Tilbud"], 3, excerpt)).toBeUndefined();
      // Uden uddrag (et tomt afsnit) gælder nummeret.
      expect(locateParagraph(doc, 2, "")).toBe(2);
      expect(locateParagraph(doc, 9, "")).toBeUndefined();
    });
  });

  it("markører fjernes fra tekst, der sættes ind i dokumentet", () => {
    expect(withoutParagraphMarks("Prisen er 499 kr. [¶3] Fristen er fem dage [¶7, ¶8].")).toBe("Prisen er 499 kr. Fristen er fem dage.");
    expect(withoutParagraphMarks("Se [1] og [note].")).toBe("Se [1] og [note].");
  });
});

describe("targetOf (#22): stedet, en handling i planen gælder", () => {
  const plan = (actions: unknown[]) => AgentPlan.parse({ summary: "s", actions }).actions;
  it("Excel: adresse, tabel og pivottabellens kilde", () => {
    const ctx = { host: "excel", text: "" } as DocContext;
    const [a, b, c, d] = plan([
      { action: "setValues", address: "Salg!$D$2", values: [[1]], reason: "" },
      { action: "appendRows", table: "Salg", values: [[1]], reason: "" },
      { action: "addPivotTable", source: "Salg", rows: ["Region"], values: [{ field: "Omsætning" }], destination: "F1", reason: "" },
      { action: "freezeRows", rows: 1, reason: "" },
    ]);
    expect([a, b, c, d].map((x) => targetOf(x, ctx))).toEqual([{ address: "Salg!D2" }, { table: "Salg" }, { table: "Salg" }, undefined]);
  });

  it("PowerPoint og Word", () => {
    const ppt = {
      host: "powerpoint",
      text: "",
      slides: [
        { layout: "Titel", shapes: [] },
        { layout: "Punkter", shapes: [{ name: "Tekst", kind: "body", text: "Pris 499 kr" }] },
      ],
    } as unknown as DocContext;
    const [s1, s2] = plan([
      { action: "setText", shape: "2.1", text: "x", reason: "" },
      { action: "replace", find: "499 kr", replace: "599 kr", reason: "" },
    ]);
    expect(targetOf(s1, ppt)).toEqual({ slide: 2 });
    expect(targetOf(s2, ppt)).toEqual({ slide: 2 });

    const word = { host: "word", text: "", isWholeDocument: true, paragraphs: [{ text: "A", style: "Normal" }] } as unknown as DocContext;
    const [w1, w2, w3] = plan([
      { action: "replace", find: "A", replace: "B", reason: "" },
      { action: "setStyle", paragraph: 1, style: "heading1", reason: "" },
      { action: "setStyle", paragraph: 9, style: "heading1", reason: "" },
    ]);
    expect([w1, w2, w3].map((x) => targetOf(x, word))).toEqual([{ find: "A" }, { paragraph: 1 }, undefined]);
    // For en markering kan afsnittet ikke findes igen med sikkerhed.
    expect(targetOf(w2, { ...word, isWholeDocument: false })).toBeUndefined();
  });
});
