import { useEffect, useState } from "react";
import { createDarkTheme, createLightTheme, tokens, type BrandVariants, type Theme } from "@fluentui/react-components";

/**
 * MIAI's designfundament (#17). Det eneste sted i task pane, hvor farver og faste mål defineres –
 * komponenterne bruger kun Fluent-tokens og konstanterne herfra (håndhæves af ESLint).
 *
 * Brandfarven er en indigo-violet rampe (OKLCH, jævne trin). Den adskiller sig fra Words blå,
 * Excels grønne og PowerPoints orange, så MIAI har sin egen identitet uden at støde mod værtsappen.
 */
export const brand: BrandVariants = {
  10: "#080717",
  20: "#120D35",
  30: "#1D1452",
  40: "#291C6F",
  50: "#35278B",
  60: "#4134A6",
  70: "#4E43BE",
  80: "#5B53D3",
  90: "#6966E5",
  100: "#7779F3",
  110: "#878DFE",
  120: "#99A2FF",
  130: "#ACB6FF",
  140: "#C7CDF1",
  150: "#DCE0F5",
  160: "#EFF3FF",
};

export const lightTheme: Theme = {
  ...createLightTheme(brand),
  // Lidt varmere, roligere flader end Fluents standard – indholdet står i kort på en let tonet baggrund.
  colorNeutralBackground2: "#F7F7FA",
  colorNeutralBackground3: "#F0F0F5",
};

export const darkTheme: Theme = {
  ...createDarkTheme(brand),
  // Brandfarvet tekst og links skal have ≥ 4,5:1 mod den mørke baggrund (trin 100 giver kun ~4:1).
  colorBrandForeground1: brand[110],
  colorBrandForeground2: brand[120],
  colorBrandForegroundLink: brand[110],
  colorBrandForegroundLinkHover: brand[120],
  colorBrandForegroundLinkPressed: brand[100],
  colorNeutralBackground1: "#262626",
  colorNeutralBackground2: "#1F1F1F",
  colorNeutralBackground3: "#2E2E2E",
};

/**
 * Faste værdier, som Fluent ikke har tokens for. Komponenter bruger disse i stedet for
 * hårdkodede pixel- og farveværdier.
 */
export const ui = {
  /** Ikonfelt foran opgaver og i headeren. */
  iconTile: "32px",
  iconTileSmall: "24px",
  /** Statusprik i headeren. */
  dot: "8px",
  /** Smalleste task pane, layoutet skal holde til. */
  minWidth: "320px",
  /** Højde på resultatets indholdsfelt, før det scroller. */
  resultMaxHeight: "45vh",
  /** Figurer (Mermaid/SVG) vises altid på hvidt "papir", så de også kan læses i mørkt tema. */
  paper: "#FFFFFF",
  paperInk: "#1B1B1F",
  /** Varighed for diskrete overgange (respekterer prefers-reduced-motion i komponenterne). */
  motion: "160ms",
  /** Hvor langt indhold glider ind ved indtoning. */
  appearOffset: "4px",
  /** Største bredde på et filnavn i en vedhæftning, før det afkortes. */
  chipNameMax: "140px",
  /** Højde på agentens kodevisning (kodetilstand), før den scroller. */
  codeMaxHeight: "240px",
  /** Mindste bredde på slide-miniaturer i designforhåndsvisningen (to kolonner i et bredt panel). */
  thumbMin: "150px",
  /** Mindste kolonnebredde i formularer med to kolonner (bryder til én kolonne i et smalt panel). */
  colMin: "130px",
  /** Brudpunkter for task pane. */
  mqNarrow: "@media (max-width: 479px)",
  mqTiny: "@media (max-width: 299px)",
} as const;

/** Opgavekategorier og deres farver (afledt af Fluents palet, så de virker i lyst og mørkt tema). */
export type TaskCategory = "write" | "present" | "analyze" | "agent" | "custom";

export const categoryColors: Record<TaskCategory, { bg: string; fg: string }> = {
  write: { bg: tokens.colorBrandBackground2, fg: tokens.colorBrandForeground1 },
  present: { bg: tokens.colorPalettePeachBackground2, fg: tokens.colorPalettePeachForeground2 },
  analyze: { bg: tokens.colorPaletteTealBackground2, fg: tokens.colorPaletteTealForeground2 },
  agent: { bg: tokens.colorPaletteCornflowerBackground2, fg: tokens.colorPaletteCornflowerForeground2 },
  custom: { bg: tokens.colorPaletteLavenderBackground2, fg: tokens.colorPaletteLavenderForeground2 },
};

/** Statusfarver (forbindelse m.m.). */
export const statusColors = {
  ok: tokens.colorPaletteGreenForeground1,
  warn: tokens.colorPaletteDarkOrangeForeground1,
  down: tokens.colorPaletteRedForeground1,
} as const;

/** Er Office' tema mørkt? isDarkTheme findes i nyere Office; ellers vurderes baggrundsfarvens lyshed. */
export function isDarkOfficeTheme(theme: { bodyBackgroundColor?: string; isDarkTheme?: boolean } | undefined): boolean {
  if (typeof theme?.isDarkTheme === "boolean") return theme.isDarkTheme;
  const bg = theme?.bodyBackgroundColor;
  if (!bg || !/^#?[0-9a-f]{6}$/i.test(bg)) return typeof window !== "undefined" && (window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false);
  const n = parseInt(bg.replace("#", ""), 16);
  return 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255) < 128;
}

/**
 * Følger Office' tema – også når brugeren skifter tema, mens panelet er åbent
 * (OfficeThemeChanged, hvor værten understøtter det; ellers systemets farveskema).
 */
export function useOfficeTheme(): Theme {
  const [dark, setDark] = useState(() => isDarkOfficeTheme(Office.context.officeTheme as never));

  useEffect(() => {
    const update = () => setDark(isDarkOfficeTheme(Office.context.officeTheme as never));
    try {
      Office.context.document?.addHandlerAsync?.(Office.EventType.OfficeThemeChanged, update);
    } catch {
      // Værten understøtter ikke temahændelser.
    }
    const mq = window.matchMedia?.("(prefers-color-scheme: dark)");
    mq?.addEventListener?.("change", update);
    return () => mq?.removeEventListener?.("change", update);
  }, []);

  return dark ? darkTheme : lightTheme;
}
