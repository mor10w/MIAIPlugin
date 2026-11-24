import { describe, expect, it } from "vitest";
import { EXCEL_ADDRESS, splitAddress } from "./excelAddress";

describe("splitAddress", () => {
  it("deler arknavn og område", () => {
    expect(splitAddress("Salg!A1:B2")).toEqual({ sheet: "Salg", range: "A1:B2" });
    expect(splitAddress("'Mit ark'!A1:B2")).toEqual({ sheet: "Mit ark", range: "A1:B2" });
    expect(splitAddress("'Bob''s'!A1")).toEqual({ sheet: "Bob's", range: "A1" });
    expect(splitAddress("Salg!$D$2:$E$5")).toEqual({ sheet: "Salg", range: "$D$2:$E$5" });
  });

  it("uden arknavn: tomt ark og hele adressen som område", () => {
    expect(splitAddress("A1:B2")).toEqual({ sheet: "", range: "A1:B2" });
    expect(splitAddress("$C$2:$C$5")).toEqual({ sheet: "", range: "$C$2:$C$5" });
  });
});

describe("EXCEL_ADDRESS", () => {
  it("matcher adresser med og uden arknavn", () => {
    expect(EXCEL_ADDRESS.test("A1:B2")).toBe(true);
    expect(EXCEL_ADDRESS.test("Salg!A1:B2")).toBe(true);
    expect(EXCEL_ADDRESS.test("'Mit ark'!A1:B2")).toBe(true);
    expect(EXCEL_ADDRESS.test("$C$2:$C$5")).toBe(true);
    expect(EXCEL_ADDRESS.test("A4-papir")).toBe(false);
    expect(EXCEL_ADDRESS.test("SUM(C2:C5)")).toBe(false);
  });

  it("arknavn med apostrof skrives med to apostroffer (#59)", () => {
    expect(EXCEL_ADDRESS.test("'Bob''s'!A1")).toBe(true);
    expect(EXCEL_ADDRESS.test("'Q1 ''26'''!A1:B2")).toBe(true);
    expect(EXCEL_ADDRESS.test("'Bob's'!A1")).toBe(false);
  });
});
