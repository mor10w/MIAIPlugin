import { describe, expect, it } from "vitest";
import type { Theme } from "@fluentui/react-components";
import { categoryColors, darkTheme, isDarkOfficeTheme, lightTheme, statusColors } from "./theme";

/** WCAG-kontrast for paletten (#17): ≥ 4,5:1 for tekst, ≥ 3:1 for grafiske UI-elementer. */

function luminance(hex: string): number {
  const n = parseInt(hex.replace("#", ""), 16);
  const ch = (v: number) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * ch((n >> 16) & 255) + 0.7152 * ch((n >> 8) & 255) + 0.0722 * ch(n & 255);
}
export function contrast(a: string, b: string): number {
  const [x, y] = [luminance(a), luminance(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}
/** tokens.* er CSS-variabler ("var(--navn)"); slå værdien op i temaet. */
const resolve = (theme: Theme, value: string) => (value.startsWith("var(--") ? (theme as unknown as Record<string, string>)[value.slice(6, -1)] : value);

for (const [name, theme] of [
  ["lyst", lightTheme],
  ["mørkt", darkTheme],
] as const) {
  describe(`kontrast – ${name} tema`, () => {
    const bg1 = theme.colorNeutralBackground1;
    const bg2 = theme.colorNeutralBackground2;
    const text: [string, string, string][] = [
      ["brødtekst på kort", theme.colorNeutralForeground1, bg1],
      ["brødtekst på baggrund", theme.colorNeutralForeground1, bg2],
      ["sekundær tekst på kort", theme.colorNeutralForeground2, bg1],
      ["dæmpet tekst på kort", theme.colorNeutralForeground3, bg1],
      ["tekst på primær knap", theme.colorNeutralForegroundOnBrand, theme.colorBrandBackground],
      ["tekst på primær knap (hover)", theme.colorNeutralForegroundOnBrand, theme.colorBrandBackgroundHover],
      ["brandfarvet tekst", theme.colorBrandForeground1, bg1],
      ["link", theme.colorBrandForegroundLink, bg1],
    ];
    for (const [what, fg, bg] of text) {
      it(`${what} ≥ 4,5:1`, () => expect(contrast(fg, bg)).toBeGreaterThanOrEqual(4.5));
    }

    for (const [cat, c] of Object.entries(categoryColors)) {
      it(`kategori ${cat}: ikon på ikonfelt ≥ 3:1`, () => expect(contrast(resolve(theme, c.fg), resolve(theme, c.bg))).toBeGreaterThanOrEqual(3));
    }
    for (const [status, c] of Object.entries(statusColors)) {
      it(`status ${status} på kort ≥ 3:1`, () => expect(contrast(resolve(theme, c), bg1)).toBeGreaterThanOrEqual(3));
    }
    it("fokusramme ≥ 3:1", () => expect(contrast(theme.colorStrokeFocus2, bg1)).toBeGreaterThanOrEqual(3));
  });
}

describe("Office-tema", () => {
  it("bruger isDarkTheme, når Office oplyser det", () => {
    expect(isDarkOfficeTheme({ isDarkTheme: true, bodyBackgroundColor: "#ffffff" })).toBe(true);
  });
  it("ellers baggrundsfarvens lyshed", () => {
    expect(isDarkOfficeTheme({ bodyBackgroundColor: "#262626" })).toBe(true);
    expect(isDarkOfficeTheme({ bodyBackgroundColor: "#F3F3F3" })).toBe(false);
  });
});
