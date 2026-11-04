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
    for (const p of ["Add a conclusion to the document", "Please edit the pricing section", "Remove the outdated appendix"]) {
      expect(looksLikeEditRequest(p), p).toBe(true);
    }
  });

  it("returnerer false på ren research og almindelig tale", () => {
    for (const p of ["Laver du research om konkurrenterne?", "Fortæl mig mere om markedet", "Hvad er den bedste model til danske tekster?"]) {
      expect(looksLikeEditRequest(p), p).toBe(false);
    }
  });

  it("æ, ø og å i starten af ordet og danske vendinger (regression: \\b kender kun a–z)", () => {
    for (const p of [
      "Ændr overskriften til Resumé",
      "Ændre tonen i afsnittet",
      "Opdatér prislisten",
      "Gør teksten kortere",
      "Gør afsnittet til en overskrift",
      "Skriv en konklusion til dokumentet",
      "Lav afsnit 2 om til punktform",
      "Ret 'kunde' til 'klient' overalt",
      "Make the intro shorter",
      "Change the title to Q3",
    ]) {
      expect(looksLikeEditRequest(p), p).toBe(true);
    }
  });

  it("tvetydige ord udløser ikke forslaget", () => {
    for (const p of [
      "Hvad mangler vi for at vinde markedet?",
      "Det er ret godt skrevet, ikke?",
      "Hvad er din email address?",
      "Hvad betyder climate change for os?",
      "What is the latest update on the project?",
      "Hvad er en tilføjelse til en kontrakt?",
    ]) {
      expect(looksLikeEditRequest(p), p).toBe(false);
    }
  });

  it("formatering og oprydning (regression: 'Gør tabellen pæn' gav kun en vejledning)", () => {
    for (const p of [
      "Gør tabellen du nævner pæn, med en moderne layout",
      "Gør arket mere overskueligt",
      "Giv tabellen et moderne look",
      "Formatér tallene som kroner",
      "Sortér tabellen efter dato",
      "Fjern de tomme rækker i tabellen",
      "Ryd op i arket",
      "Format the table as currency",
      "Make the table nicer",
      "Clean up the sheet",
    ]) {
      expect(looksLikeEditRequest(p), p).toBe(true);
    }
    for (const p of ["Hvilket format bruger I til datoer?", "What format should dates use?", "Hvad er et pænt resultat for kvartalet?"]) {
      expect(looksLikeEditRequest(p), p).toBe(false);
    }
  });

  it("returnerer false på tom tekst", () => {
    expect(looksLikeEditRequest("")).toBe(false);
  });
});
