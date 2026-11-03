import { z } from "zod";

/**
 * Design af en hel præsentation (PowerPoint, ADR-0014): modellen læser indholdet og vælger
 * stil, farver, skrifttyper og et layout pr. slide inden for faste rammer. Add-in'et bygger en
 * slidemaster med navngivne layouts ud fra valget, så hele præsentationen følger samme design.
 *
 * Modellen tegner ikke selv: geometrien ligger i faste skabeloner, og farverne tjekkes for
 * kontrast (WCAG) – så resultatet altid er læsbart og konsistent.
 */

/** Skrifttyper, der følger med Office på Windows og Mac (eller erstattes pænt). */
export const SAFE_FONTS = [
  "Aptos",
  "Aptos Display",
  "Calibri",
  "Segoe UI",
  "Arial",
  "Helvetica Neue",
  "Verdana",
  "Trebuchet MS",
  "Georgia",
  "Cambria",
  "Garamond",
] as const;
export const DesignFont = z.enum(SAFE_FONTS);

/** Hex-farve; normaliseres til "#RRGGBB" i fixPalette (en transform kan ikke stå i JSON Schema). */
const hex = z.string().regex(/^#?[0-9a-fA-F]{6}$/);
const norm = (h: string) => `#${h.replace("#", "").toUpperCase()}`;

export const DesignStyle = z
  .enum(["clean", "bold", "editorial"])
  .describe(
    "clean = light background, accent bar, lots of white space; bold = dark title/section slides with a coloured band; editorial = serif headings, thin rules, calm",
  );
export type DesignStyle = z.infer<typeof DesignStyle>;

const bullets = z.array(z.string().max(160)).min(1).max(8);

export const DesignSlide = z.discriminatedUnion("layout", [
  z.object({ layout: z.literal("title"), title: z.string().max(120), subtitle: z.string().max(200).optional() }),
  z.object({ layout: z.literal("section"), title: z.string().max(120), subtitle: z.string().max(200).optional() }),
  z.object({ layout: z.literal("bullets"), title: z.string().max(120), bullets }),
  z.object({
    layout: z.literal("twoColumn"),
    title: z.string().max(120),
    left: z.object({ heading: z.string().max(80).optional(), bullets }),
    right: z.object({ heading: z.string().max(80).optional(), bullets }),
  }),
  z.object({ layout: z.literal("quote"), quote: z.string().max(300), attribution: z.string().max(120).optional() }),
  z.object({
    layout: z.literal("kpi"),
    title: z.string().max(120),
    kpis: z
      .array(z.object({ value: z.string().max(16), label: z.string().max(60) }))
      .min(1)
      .max(4),
  }),
  z.object({ layout: z.literal("closing"), title: z.string().max(120), subtitle: z.string().max(200).optional() }),
]);
export type DesignSlide = z.infer<typeof DesignSlide>;
export type DesignLayout = DesignSlide["layout"];

export const DesignPlan = z.object({
  name: z.string().max(60).describe("Short name for the design, e.g. 'Nordisk klar'"),
  rationale: z.string().max(400).describe("One or two sentences: why this design fits the content, in the user's language"),
  style: DesignStyle,
  palette: z
    .object({
      background: hex.describe("Slide background"),
      text: hex.describe("Main text on the background"),
      muted: hex.describe("Secondary text (subtitles, captions)"),
      primary: hex.describe("Brand/primary colour: bands, title slide background in 'bold'"),
      accent: hex.describe("Highlights: bars, numbers, bullets"),
    })
    .describe("Colours chosen from the content's mood and subject; restrained, max two strong colours"),
  fonts: z.object({ heading: DesignFont, body: DesignFont }),
  slides: z.array(DesignSlide).min(1).max(40),
});
export type DesignPlan = z.infer<typeof DesignPlan>;

// --- Kontrast (WCAG 2.1) -----------------------------------------------------------------

const lum = (h: string) => {
  const n = parseInt(h.replace("#", ""), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
};

export function contrastRatio(a: string, b: string): number {
  const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}

/** Blander en farve mod sort (t<0) eller hvid (t>0) i små trin. */
function shade(h: string, t: number): string {
  const n = parseInt(h.replace("#", ""), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(t > 0 ? v + (255 - v) * t : v * (1 + t)));
  return `#${ch
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

/** Gør `fg` mørkere eller lysere, indtil kontrasten mod `bg` er mindst `min`. */
function ensure(fg: string, bg: string, min: number): string {
  if (contrastRatio(fg, bg) >= min) return fg;
  const dir = lum(bg) > 0.4 ? -1 : 1;
  for (let i = 1; i <= 20; i++) {
    const c = shade(fg, dir * i * 0.05);
    if (contrastRatio(c, bg) >= min) return c;
  }
  return dir < 0 ? "#111111" : "#FFFFFF";
}

/** Tekstfarve (hvid eller næsten sort) med bedst kontrast på en flade, fx på primærfarven. */
export const onColor = (bg: string) => (contrastRatio("#FFFFFF", bg) >= contrastRatio("#1A1A1A", bg) ? "#FFFFFF" : "#1A1A1A");

/**
 * Retter paletten, så den er læsbar: brødtekst ≥ 4,5:1 mod baggrunden, sekundær tekst og
 * accent ≥ 3:1 (store tal og streger). Modellens farver bevares, når de holder.
 */
export function fixPalette(raw: DesignPlan["palette"]): DesignPlan["palette"] {
  const p = Object.fromEntries(Object.entries(raw).map(([k, v]) => [k, norm(v)])) as DesignPlan["palette"];
  return {
    background: p.background,
    text: ensure(p.text, p.background, 4.5),
    muted: ensure(p.muted, p.background, 4.5),
    primary: p.primary,
    accent: ensure(p.accent, p.background, 3),
  };
}

/** Layoutnavne i masteren – også dem, agenten bagefter kan oprette nye slides fra. */
export const DESIGN_LAYOUT_NAMES: Record<DesignLayout, { da: string; en: string }> = {
  title: { da: "Titel", en: "Title" },
  section: { da: "Sektion", en: "Section" },
  bullets: { da: "Punkter", en: "Bullets" },
  twoColumn: { da: "To kolonner", en: "Two columns" },
  quote: { da: "Citat", en: "Quote" },
  kpi: { da: "Nøgletal", en: "Key figures" },
  closing: { da: "Afslutning", en: "Closing" },
};

/** Retningslinjer til modellen: design ud fra indholdet, et layout pr. slide, korte tekster. */
export const DESIGN_GUIDE = [
  "You are a presentation designer. Read the content and design ONE consistent look for the whole deck, then lay out the content slide by slide.",
  "Choose the style and palette from the subject and audience (e.g. finance → calm blue/green, sustainability → greens, creative → one bold accent). Use a restrained palette: one primary, one accent, neutral background and text. Do not pick colours that clash or vibrate.",
  `Fonts: choose a heading and a body font from: ${SAFE_FONTS.join(", ")}. Serif headings (Georgia, Cambria, Garamond) suit 'editorial'; sans-serif suits 'clean' and 'bold'.`,
  "Layouts: title (first slide), section (divides parts), bullets (3-6 short bullets), twoColumn (comparisons, before/after, pros/cons), quote (one strong statement), kpi (1-4 big numbers with labels), closing (last slide: thanks/next steps/contact).",
  "Rules: one idea per slide; max 6 bullets of max ~10 words; turn numbers into kpi slides; split long content over several slides; keep the user's facts and language; do not invent figures.",
].join("\n");
