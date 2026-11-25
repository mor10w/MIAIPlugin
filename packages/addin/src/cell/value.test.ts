import { describe, expect, it } from "vitest";
import { cleanAnswer, inputToText, MAX_CELL_CHARS, toCellValue } from "./value";

describe("fra Excels argumenter til tekst", () => {
  it("én celle er ren tekst, et område er TSV", () => {
    expect(inputToText([["Aarhus"]])).toBe("Aarhus");
    expect(inputToText([[42]])).toBe("42");
    expect(inputToText("direkte tekst")).toBe("direkte tekst");
    expect(inputToText(null)).toBe("");
    expect(inputToText([[null]])).toBe("");
    expect(
      inputToText([
        ["Navn", "By", 2300],
        ["Acme A/S", null, true],
      ]),
    ).toBe("Navn\tBy\t2300\nAcme A/S\t\ttrue");
  });
  it("tabulatorer og linjeskift i en celle ødelægger ikke TSV'en, og tomme rækker i bunden fjernes", () => {
    expect(
      inputToText([
        ["a\tb", "to\nlinjer"],
        ["", ""],
        [null, null],
      ]),
    ).toBe("a b\tto linjer");
  });
});

describe("svaret ryddes op", () => {
  it("fjerner kodehegn og omsluttende citationstegn, men ikke citationstegn inde i teksten", () => {
    expect(cleanAnswer("  Danmark \n")).toBe("Danmark");
    expect(cleanAnswer("```\nDanmark\n```")).toBe("Danmark");
    expect(cleanAnswer('"Danmark"')).toBe("Danmark");
    expect(cleanAnswer('Han sagde "ja" og gik')).toBe('Han sagde "ja" og gik');
    expect(cleanAnswer('"ja" eller "nej"')).toBe('"ja" eller "nej"');
  });
  it("afkorter ved cellens grænse", () => {
    expect(cleanAnswer("x".repeat(40_000))).toHaveLength(MAX_CELL_CHARS);
  });
});

describe("tal som tal, kun når det er entydigt", () => {
  it("hele tal og decimaltal på dansk og engelsk", () => {
    expect(toCellValue("42")).toBe(42);
    expect(toCellValue("-7")).toBe(-7);
    expect(toCellValue("1234,5")).toBe(1234.5);
    expect(toCellValue("1234.5")).toBe(1234.5);
    expect(toCellValue("0,125")).toBe(0.125);
    expect(toCellValue("1.234,56")).toBe(1234.56);
    expect(toCellValue("1,234.56")).toBe(1234.56);
    expect(toCellValue("1.234.567")).toBe(1234567);
    expect(toCellValue("1234,567")).toBe(1234.567);
  });
  it("foranstillet nul, tvetydige tal og tal med tekst forbliver tekst", () => {
    for (const s of ["0045", "007", "1.234", "1,234", "12.345", "2300 København", "45 %", "kr. 100", "1.2.3", "1,23,4", "12.34.567", "+4512345678", "", "."])
      expect(toCellValue(s)).toBe(s);
  });
  it("meget lange cifferrækker (kortnumre) forbliver tekst", () => {
    expect(toCellValue("1234567890123456")).toBe("1234567890123456");
    expect(toCellValue("123456789012345")).toBe(123456789012345);
  });
  it("et svar, der ligner en formel, er tekst", () => {
    expect(toCellValue("=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
  });
});
