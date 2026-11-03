import { DESIGN_LAYOUT_NAMES, type DesignLayout, type DesignPlan, type DesignSlide } from "@miai/shared";
import { kpiBoxes, layoutName, resolvePalette, templateFor, type Palette, type TemplatePlaceholder } from "./templates";

/**
 * Bygger en .pptx med designets slidemaster og layouts og slides med indholdet (ADR-0014).
 * Filen sættes ind i den åbne præsentation med insertSlidesFromBase64 og "behold
 * kildeformatering", så masteren og layoutene følger med – PowerPoints JavaScript-API kan ikke
 * oprette mastere direkte.
 */
export async function buildDeck(plan: DesignPlan, lang: "da" | "en"): Promise<string> {
  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = plan.name;
  pptx.theme = { headFontFace: plan.fonts.heading, bodyFontFace: plan.fonts.body };
  const pal = resolvePalette(plan);
  const hex = (c: string) => c.replace("#", "");
  const font = (f: "heading" | "body") => (f === "heading" ? plan.fonts.heading : plan.fonts.body);

  // Ét layout pr. slidetype, der bruges i planen (+ punkter, så agenten altid kan lave indholdsslides).
  const used = [...new Set<DesignLayout>([...plan.slides.map((s) => s.layout), "bullets"])];
  const names = new Map<DesignLayout, string>();
  for (const layout of used) {
    const t = templateFor(plan.style, layout);
    const name = layoutName(plan, DESIGN_LAYOUT_NAMES[layout][lang]);
    names.set(layout, name);
    pptx.defineSlideMaster({
      title: name,
      background: { color: hex(pal[t.background]) },
      objects: [
        ...t.shapes.map((r) => ({ rect: { x: r.x, y: r.y, w: r.w, h: r.h, fill: { color: hex(pal[r.fill]) }, line: { type: "none" as const } } })),
        ...t.placeholders.map((p) => ({ placeholder: { options: phOptions(p, pal, font), text: "" } })),
      ],
    });
  }

  for (const s of plan.slides) {
    const slide = pptx.addSlide({ masterName: names.get(s.layout)! });
    fill(slide, s, pal, font);
  }
  return (await pptx.write({ outputType: "base64" })) as string;
}

function phOptions(p: TemplatePlaceholder, pal: Palette, font: (f: "heading" | "body") => string) {
  return {
    name: p.name,
    type: p.type,
    x: p.x,
    y: p.y,
    w: p.w,
    h: p.h,
    fontFace: font(p.font),
    fontSize: p.size,
    color: pal[p.color].replace("#", ""),
    bold: !!p.bold,
    italic: !!p.italic,
    align: p.align ?? "left",
    valign: p.valign ?? "top",
    margin: 0,
  };
}

type Slide = ReturnType<InstanceType<typeof import("pptxgenjs").default>["addSlide"]>;

/** Indholdet i layoutets pladsholdere; nøgletal som tekst placeret efter antal. */
function fill(slide: Slide, s: DesignSlide, pal: Palette, font: (f: "heading" | "body") => string) {
  const text = (ph: string, t: string | undefined) => t && slide.addText(t, { placeholder: ph });
  const list = (ph: string, items: string[]) =>
    slide.addText(
      items.map((b) => ({ text: b, options: { bullet: { indent: 18 }, breakLine: true } })),
      { placeholder: ph },
    );
  switch (s.layout) {
    case "title":
    case "section":
    case "closing":
      text("title", s.title);
      text("subtitle", s.subtitle);
      break;
    case "bullets":
      text("title", s.title);
      list("body", s.bullets);
      break;
    case "twoColumn":
      text("title", s.title);
      text("leftHead", s.left.heading);
      list("left", s.left.bullets);
      text("rightHead", s.right.heading);
      list("right", s.right.bullets);
      break;
    case "quote":
      text("quote", `“${s.quote.replace(/^["“]|["”]$/g, "")}”`);
      text("attribution", s.attribution && `— ${s.attribution}`);
      break;
    case "kpi": {
      text("title", s.title);
      kpiBoxes(s.kpis.length).forEach((b, i) => {
        const k = s.kpis[i];
        slide.addText(k.value, {
          x: b.x,
          y: b.valueY,
          w: b.w,
          h: 1.4,
          fontFace: font("heading"),
          fontSize: s.kpis.length > 3 ? 44 : 54,
          bold: true,
          color: pal.accent.replace("#", ""),
          valign: "bottom",
          margin: 0,
        });
        slide.addText(k.label, {
          x: b.x,
          y: b.labelY,
          w: b.w,
          h: 0.8,
          fontFace: font("body"),
          fontSize: 18,
          color: pal.muted.replace("#", ""),
          valign: "top",
          margin: 0,
        });
      });
      break;
    }
  }
}
