import { describe, expect, it } from "vitest";
import { looksLikeEditRequest } from "./editIntent.js";

describe("looksLikeEditRequest", () => {
  it("genkender danske ændringsord med fleksion", () => {
    for (const p of [
      "Tilføj et afsnit om leveringsbetingelser",
      "Kan du indsætte en overskrift i toppen?",
      "Ret alle forekomster af 'kunde' til 'klient'",
      "Slet den sidste sektion",
      "Opdater prislisten i dokumentet",
      "Dokumentet mangler en konklusion – skriv den",
    ]) {
      expect(looksLikeEditRequest(p), p).toBe(true);
    }
  });

  it("genkender engelske ændringsord med fleksion", () => {
    for (const p of [
      "Add a conclusion to the document",
      "Please edit the pricing section",
      "Remove the outdated appendix",
    ]) {
      expect(looksLikeEditRequest(p), p).toBe(true);
    }
  });

  it("returnerer false på ren research og almindelig tale", () => {
    for (const p of [
      "Laver du research om konkurrenterne?",
      "Fortæl mig mere om markedet",
      "Hvad er den bedste model til danske tekster?",
    ]) {
      expect(looksLikeEditRequest(p), p).toBe(false);
    }
  });

  it("returnerer false på tom tekst", () => {
    expect(looksLikeEditRequest("")).toBe(false);
  });
});
