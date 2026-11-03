import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { contrastRatio, DesignPlan } from "@miai/shared";
import { buildDeck } from "./buildDeck";
import { resolvePalette, templateFor } from "./templates";

const plan = DesignPlan.parse({
  name: "Nordisk klar",
  rationale: "Rolig til bestyrelsen.",
  style: "clean",
  palette: { background: "#FFFFFF", text: "#1B2A3A", muted: "#C8CDD2", primary: "#0A5C8A", accent: "#E07A1F" },
  fonts: { heading: "Georgia", body: "Aptos" },
  slides: [
    { layout: "title", title: "Kvartalsrapport Q3", subtitle: "Bestyrelsesmøde" },
    {
      layout: "kpi",
      title: "Nøgletal",
      kpis: [
        { value: "+12 %", label: "Omsætning" },
        { value: "38 %", label: "Margin" },
      ],
    },
    { layout: "twoColumn", title: "Før og nu", left: { heading: "Før", bullets: ["Manuelt"] }, right: { heading: "Nu", bullets: ["Automatisk"] } },
    { layout: "quote", quote: "Det bedste kvartal nogensinde", attribution: "CFO" },
    { layout: "closing", title: "Tak" },
  ],
});

describe("pptx med designets master (ADR-0014)", () => {
  it("indeholder én master med navngivne layouts, pladsholdere, skrifttyper og farver, og ét slide pr. planslide", async () => {
    const zip = await JSZip.loadAsync(Buffer.from(await buildDeck(plan, "da"), "base64"));
    const files = Object.keys(zip.files);
    expect(files.filter((f) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f))).toHaveLength(1);
    expect(files.filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))).toHaveLength(5);

    const layouts = await Promise.all(files.filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f)).map((f) => zip.file(f)!.async("string")));
    const names = layouts.map((x) => x.match(/<p:cSld name="([^"]+)"/)?.[1]);
    // Et layout pr. brugt slidetype + "Punkter", så agenten altid kan oprette indholdsslides.
    for (const n of ["Titel", "Nøgletal", "To kolonner", "Citat", "Afslutning", "Punkter"]) expect(names).toContain(`Nordisk klar – ${n}`);

    const punkter = layouts[names.indexOf("Nordisk klar – Punkter")];
    expect(punkter).toMatch(/<p:ph[^>]*type="title"/);
    expect(punkter).toMatch(/<p:ph[^>]*type="body"/);
    expect(punkter).toContain('typeface="Georgia"');
    expect(punkter).toContain("E07A1F"); // accentstregen under titlen

    // Indholdet står i slides, i layoutets pladsholdere.
    const slide3 = await zip.file("ppt/slides/slide3.xml")!.async("string");
    expect(slide3).toContain("Før og nu");
    expect(slide3).toContain("Automatisk");
    const theme = await zip.file("ppt/theme/theme1.xml")!.async("string");
    expect(theme).toMatch(/<a:latin typeface="Georgia"/);
  });

  it("paletten rettes for kontrast, før den bruges (for lys sekundærtekst bliver mørkere)", () => {
    const pal = resolvePalette(plan);
    expect(contrastRatio(pal.muted, pal.background)).toBeGreaterThanOrEqual(4.5);
    expect(pal.onPrimary).toBe("#FFFFFF");
  });

  it("alle stilarter har skabeloner for alle layouts, inden for slidets mål", () => {
    for (const style of ["clean", "bold", "editorial"] as const)
      for (const layout of ["title", "section", "bullets", "twoColumn", "quote", "kpi", "closing"] as const) {
        const t = templateFor(style, layout);
        for (const p of [...t.placeholders, ...t.shapes]) {
          expect(p.x + p.w, `${style}/${layout}`).toBeLessThanOrEqual(13.34);
          expect(p.y + p.h, `${style}/${layout}`).toBeLessThanOrEqual(7.5);
        }
        expect(t.placeholders.some((p) => p.name === "title" || p.name === "quote")).toBe(true);
      }
  });
});
