import { describe, expect, it } from "vitest";
import { z } from "zod";
import { contrastRatio, DesignPlan, fixPalette, onColor } from "./design.js";
import { TASKS } from "./tasks.js";
import { TaskRequest } from "./types.js";

const plan = {
  name: "Nordisk klar",
  rationale: "Rolig og troværdig til et finansielt publikum.",
  style: "clean",
  palette: { background: "#FFFFFF", text: "#1B2A3A", muted: "#5B6B7B", primary: "#0A5C8A", accent: "#E07A1F" },
  fonts: { heading: "Aptos Display", body: "Aptos" },
  slides: [
    { layout: "title", title: "Kvartalsrapport Q3", subtitle: "Bestyrelsesmøde" },
    { layout: "kpi", title: "Nøgletal", kpis: [{ value: "+12 %", label: "Omsætning" }] },
    { layout: "bullets", title: "Fokus", bullets: ["Norge", "Webshop i Sverige"] },
    { layout: "closing", title: "Tak" },
  ],
};

describe("design af præsentation (ADR-0014)", () => {
  it("skemaet kan bruges som json_schema og normaliserer farver", () => {
    expect(JSON.stringify(z.toJSONSchema(DesignPlan))).toContain('"twoColumn"');
    const p = DesignPlan.parse({ ...plan, palette: { ...plan.palette, primary: "0a5c8a" } });
    expect(fixPalette(p.palette).primary).toBe("#0A5C8A");
  });
  it("afviser skrifttyper uden for Office og ukendte layouts", () => {
    expect(DesignPlan.safeParse({ ...plan, fonts: { heading: "Comic Sans MS", body: "Aptos" } }).success).toBe(false);
    expect(DesignPlan.safeParse({ ...plan, slides: [{ layout: "collage", title: "x" }] }).success).toBe(false);
  });
  it("fixPalette sikrer læsbar kontrast og bevarer farver, der holder", () => {
    const ok = fixPalette(DesignPlan.parse(plan).palette);
    expect(ok.text).toBe("#1B2A3A");
    const bad = fixPalette({ background: "#FFFFFF", text: "#BBBBBB", muted: "#DDDDDD", primary: "#FFD400", accent: "#FFF3B0" });
    expect(contrastRatio(bad.text, "#FFFFFF")).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(bad.muted, "#FFFFFF")).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(bad.accent, "#FFFFFF")).toBeGreaterThanOrEqual(3);
    // Mørk baggrund: lys tekst.
    const dark = fixPalette({ background: "#101820", text: "#303840", muted: "#404850", primary: "#22AA66", accent: "#333355" });
    expect(contrastRatio(dark.text, "#101820")).toBeGreaterThanOrEqual(4.5);
  });
  it("onColor vælger hvid eller sort tekst på en flade", () => {
    expect(onColor("#0A5C8A")).toBe("#FFFFFF");
    expect(onColor("#FFD400")).toBe("#1A1A1A");
  });
  it("prompten giver modellen designreglerne og præsentationens indhold", () => {
    const task = TASKS.find((t) => t.id === "design-deck")!;
    const msgs = task.buildMessages(
      TaskRequest.parse({
        context: { host: "powerpoint", text: "", slides: [{ layout: "Title Slide", shapes: [{ name: "Title 1", kind: "title", text: "Kvartalsrapport" }] }] },
        params: { prompt: "Rolig og professionel" },
      }),
    );
    expect(msgs[0].content).toMatch(/ONE consistent look for the whole deck/);
    expect(msgs[0].content).toContain("Aptos");
    expect(msgs[1].content).toContain("1.1 [title] Title 1: Kvartalsrapport");
  });
});
