import { describe, expect, it } from "vitest";
import { AgentPlan, type DocContext } from "@miai/shared";
import { findRefs, targetOf, usedRangeOf } from "./refs";

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
