import { fixPalette, onColor, type DesignLayout, type DesignPlan, type DesignStyle } from "@miai/shared";

/**
 * Designskabeloner til præsentationens master (ADR-0014). Én kilde til både den .pptx, der
 * sættes ind i PowerPoint, og forhåndsvisningen i task pane – så det, brugeren ser, er det,
 * der bliver lavet.
 *
 * Mål i tommer på et 16:9-slide (13,333 × 7,5). Farver angives som roller, der slås op i
 * den (kontrast-rettede) palet.
 */
export const SLIDE_W = 13.333;
export const SLIDE_H = 7.5;

export type ColorRole = "background" | "text" | "muted" | "primary" | "accent" | "onPrimary";

export interface TemplateRect {
  x: number;
  y: number;
  w: number;
  h: number;
  fill: ColorRole;
}

/** En pladsholder i layoutet (titel eller tekst), som PowerPoint viser og brugeren kan skrive i. */
export interface TemplatePlaceholder {
  name: "title" | "subtitle" | "body" | "left" | "right" | "leftHead" | "rightHead" | "quote" | "attribution";
  type: "title" | "body";
  x: number;
  y: number;
  w: number;
  h: number;
  font: "heading" | "body";
  size: number;
  color: ColorRole;
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center";
  valign?: "top" | "middle" | "bottom";
  bullets?: boolean;
}

export interface Template {
  background: ColorRole;
  shapes: TemplateRect[];
  placeholders: TemplatePlaceholder[];
}

export type Palette = Record<ColorRole, string>;

/** Paletten med kontrast-rettelser og tekstfarven til flader i primærfarven. */
export function resolvePalette(plan: DesignPlan): Palette {
  const p = fixPalette(plan.palette);
  return { ...p, onPrimary: onColor(p.primary) };
}

const M = 0.7; // margen
const CW = SLIDE_W - 2 * M; // indholdsbredde

/** Fælles titel på indholdsslides – placering og pynt afhænger af stilen. */
function contentChrome(style: DesignStyle): { shapes: TemplateRect[]; title: TemplatePlaceholder; top: number } {
  const title: TemplatePlaceholder = {
    name: "title",
    type: "title",
    x: M,
    y: 0.45,
    w: CW,
    h: 1.0,
    font: "heading",
    size: 32,
    color: "text",
    bold: style !== "editorial",
    valign: "bottom",
  };
  if (style === "bold") return { shapes: [{ x: 0, y: 0, w: SLIDE_W, h: 0.22, fill: "primary" }], title: { ...title, y: 0.55 }, top: 1.85 };
  if (style === "editorial") return { shapes: [{ x: M, y: 1.55, w: CW, h: 0.02, fill: "muted" }], title, top: 1.85 };
  return { shapes: [{ x: M, y: 1.55, w: 1.1, h: 0.08, fill: "accent" }], title, top: 1.9 };
}

/** Titel-, sektions- og afslutningsslides: store, rolige flader. */
function hero(style: DesignStyle, kind: "title" | "section" | "closing"): Template {
  const strong = style === "bold" || (style === "clean" && kind === "section");
  const fg: ColorRole = strong ? "onPrimary" : "text";
  const sub: ColorRole = strong ? "onPrimary" : "muted";
  const centered = style === "editorial";
  const size = kind === "section" ? 40 : 48;
  const shapes: TemplateRect[] =
    style === "clean" && !strong
      ? [{ x: M, y: 2.35, w: 0.14, h: 2.3, fill: "accent" }]
      : style === "editorial"
        ? [{ x: SLIDE_W / 2 - 0.8, y: 4.55, w: 1.6, h: 0.03, fill: "accent" }]
        : style === "bold"
          ? [{ x: 0, y: SLIDE_H - 0.6, w: SLIDE_W, h: 0.6, fill: "accent" }]
          : [];
  const x = style === "clean" && !strong ? M + 0.45 : M;
  return {
    background: strong ? "primary" : "background",
    shapes,
    placeholders: [
      {
        name: "title",
        type: "title",
        x,
        y: 2.2,
        w: CW - (x - M),
        h: 2.2,
        font: "heading",
        size,
        color: fg,
        bold: style !== "editorial",
        align: centered ? "center" : "left",
        valign: "bottom",
      },
      {
        name: "subtitle",
        type: "body",
        x,
        y: 4.75,
        w: CW - (x - M),
        h: 1.2,
        font: "body",
        size: 20,
        color: sub,
        align: centered ? "center" : "left",
        valign: "top",
      },
    ],
  };
}

export function templateFor(style: DesignStyle, layout: DesignLayout): Template {
  if (layout === "title" || layout === "section" || layout === "closing") return hero(style, layout);
  const { shapes, title, top } = contentChrome(style);
  const bodyH = SLIDE_H - top - 0.6;
  switch (layout) {
    case "bullets":
      return {
        background: "background",
        shapes,
        placeholders: [
          title,
          { name: "body", type: "body", x: M, y: top, w: CW, h: bodyH, font: "body", size: 22, color: "text", bullets: true, valign: "top" },
        ],
      };
    case "twoColumn": {
      const gap = 0.6;
      const w = (CW - gap) / 2;
      const col = (x: number, side: "left" | "right"): TemplatePlaceholder[] => [
        {
          name: side === "left" ? "leftHead" : "rightHead",
          type: "body",
          x,
          y: top,
          w,
          h: 0.6,
          font: "heading",
          size: 22,
          color: "accent",
          bold: true,
          valign: "top",
        },
        { name: side, type: "body", x, y: top + 0.7, w, h: bodyH - 0.7, font: "body", size: 20, color: "text", bullets: true, valign: "top" },
      ];
      return { background: "background", shapes, placeholders: [title, ...col(M, "left"), ...col(M + w + gap, "right")] };
    }
    case "quote":
      return {
        background: style === "bold" ? "primary" : "background",
        shapes: style === "bold" ? [] : [{ x: M, y: 1.6, w: 0.12, h: 3.6, fill: "accent" }],
        placeholders: [
          {
            name: "quote",
            type: "body",
            x: M + 0.5,
            y: 1.4,
            w: CW - 1,
            h: 3.9,
            font: "heading",
            size: 34,
            color: style === "bold" ? "onPrimary" : "text",
            italic: true,
            valign: "middle",
          },
          {
            name: "attribution",
            type: "body",
            x: M + 0.5,
            y: 5.45,
            w: CW - 1,
            h: 0.7,
            font: "body",
            size: 18,
            color: style === "bold" ? "onPrimary" : "muted",
            valign: "top",
          },
        ],
      };
    case "kpi":
      // Tallene placeres af kpiBoxes (antallet varierer); layoutet har titlen.
      return { background: "background", shapes, placeholders: [title] };
  }
}

/** Placering af 1–4 nøgletal (værdi og etiket) under titlen. */
export function kpiBoxes(n: number): { x: number; w: number; valueY: number; labelY: number }[] {
  const gap = 0.5;
  const w = (CW - gap * (n - 1)) / n;
  return Array.from({ length: n }, (_, i) => ({ x: M + i * (w + gap), w, valueY: 2.6, labelY: 4.2 }));
}

/** Layoutets navn i PowerPoint, fx "Nordisk klar – Punkter". */
export const layoutName = (plan: DesignPlan, layoutLabel: string) => `${plan.name.trim() || "MIAI"} – ${layoutLabel}`;
