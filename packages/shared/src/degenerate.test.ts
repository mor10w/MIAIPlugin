import { describe, expect, it } from "vitest";
import { isDegenerating } from "./degenerate.js";

const loop = (pattern: string, n: number) => Array(n).fill(pattern).join(" ");

describe("isDegenerating", () => {
  it("detekterer 2-ords sløjfe (side 6)", () => {
    expect(isDegenerating(`Jeg kigger på indholdet og bygger figuren. ${loop("side 6", 6)}`)).toBe(true);
  });

  it("detekterer 3-ords sløjfe (side 6 tabel)", () => {
    expect(isDegenerating(`Her er en figur baseret på indholdet: ${loop("side 6 tabel", 6)}`)).toBe(true);
  });

  it("detekterer 1-ords sløjfe først ved et langt løb", () => {
    expect(isDegenerating(`Svar: ${loop("ja", 11)}`)).toBe(false);
    expect(isDegenerating(`Svar: ${loop("ja", 12)}`)).toBe(true);
  });

  it("normal tekst triggerer ikke", () => {
    expect(isDegenerating("Ja — her er et forslag til en mere grafisk statusfigur ud fra tabellen med kolonner og rækker.")).toBe(false);
  });

  it("kort tekst triggerer ikke", () => {
    expect(isDegenerating("side 6 side 6")).toBe(false);
  });

  it("hyppige men ikke sammenhængende gentagelser triggerer ikke", () => {
    expect(isDegenerating("Status: færdig i team A, færdig i team B, færdig i team C, færdig i team D, færdig i team E, færdig i team F")).toBe(false);
  });

  it("oversat tabelkolonne med samme værdi triggerer ikke", () => {
    expect(isDegenerating(["Status", ...Array(8).fill("Done")].join("\n"))).toBe(false);
  });

  it("sløjfen opdages, selvom den kommer efter en lang tekst", () => {
    expect(isDegenerating(`${"Lorem ipsum dolor sit amet. ".repeat(500)}${loop("side 6 tabel", 6)}`)).toBe(true);
  });
});
