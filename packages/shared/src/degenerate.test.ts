import { describe, expect, it } from "vitest";
import { isDegenerating } from "./degenerate.js";

describe("isDegenerating", () => {
  it("detekterer 1-ords sløjfe (side 6)", () => {
    const text = "Jeg kigger på indholdet og bygger figuren. side 6 side 6 side 6 side 6 side 6";
    expect(isDegenerating(text)).toBe(true);
  });

  it("detekterer 2-ords sløjfe (side 6 / tabel)", () => {
    const text = "Her er en figur baseret på indholdet: side 6 tabel side 6 tabel side 6 tabel side 6 tabel side 6";
    expect(isDegenerating(text)).toBe(true);
  });

  it("normal tekst triggerer ikke", () => {
    expect(isDegenerating("Ja — her er et forslag til en mere grafisk statusfigur ud fra tabellen med kolonner og rækker.")).toBe(false);
  });

  it("kort tekst triggerer ikke", () => {
    expect(isDegenerating("side 6 side 6")).toBe(false);
  });

  it("ordinær gentagelse (fx tabel-celler) triggerer ikke ved <5 forekomster", () => {
    expect(isDegenerating("Status: færdig i team A, færdig i team B, færdig i team C og færdig i team D")).toBe(false);
  });
});
