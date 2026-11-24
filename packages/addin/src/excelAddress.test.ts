import { describe, expect, it } from "vitest";
import { EXCEL_ADDRESS } from "./excelAddress";

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
