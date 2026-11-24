import { describe, expect, it } from "vitest";
import { dictionaries, makeT } from "./i18n";

/** Placeholders i en tekst, fx "Skrev {n} celler i {address}" → ["address", "n"]. */
const placeholders = (text: string) => [...new Set([...text.matchAll(/\{(\w+)\}/g)].map((m) => m[1]))].sort();

describe("i18n (#58)", () => {
  // Nøglerne sikres af typen Dict; det, TypeScript ikke fanger, er placeholders: makeT erstatter
  // en ukendt {navn} med tom streng, så en stavefejl i den ene oversættelse giver bare en tekst uden værdien.
  it("da og en bruger de samme {placeholders} for hver nøgle", () => {
    const { da, en } = dictionaries;
    const mismatches = Object.keys(da)
      .map((key) => key as keyof typeof da)
      .filter((key) => placeholders(da[key]).join() !== placeholders(en[key]).join())
      .map((key) => `${key}: da {${placeholders(da[key]).join(", ")}} ≠ en {${placeholders(en[key]).join(", ")}}`);
    expect(mismatches).toEqual([]);
  });

  it("makeT indsætter værdier i begge sprog", () => {
    const key = (Object.keys(dictionaries.da) as (keyof typeof dictionaries.da)[]).find((k) => placeholders(dictionaries.da[k]).length > 0)!;
    const vars = Object.fromEntries(placeholders(dictionaries.da[key]).map((p) => [p, `«${p}»`]));
    for (const lang of ["da", "en"] as const) {
      const text = makeT(lang)(key, vars);
      for (const p of Object.keys(vars)) expect(text).toContain(`«${p}»`);
      expect(text).not.toMatch(/\{\w+\}/);
    }
  });
});
