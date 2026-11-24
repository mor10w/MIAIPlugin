import { describe, expect, it } from "vitest";
import { splitAddress } from "./excel";

describe("splitAddress (#57): parser af Excel-adresse", () => {
  it("parser arknavn og range fra en adresser med arknavn", () => {
    expect(splitAddress("Salg!A1:B2")).toEqual({ sheet: "Salg", range: "A1:B2" });
    expect(splitAddress("'Mit ark'!A1:B2")).toEqual({ sheet: "Mit ark", range: "A1:B2" });
    expect(splitAddress("Salg!$D$2:$E$5")).toEqual({ sheet: "Salg", range: "$D$2:$E$5" });
  });

  it("parser adresser uden arknavn som tomme sheet og helt range", () => {
    expect(splitAddress("A1:B2")).toEqual({ sheet: "", range: "A1:B2" });
    expect(splitAddress("$C$2:$C$5")).toEqual({ sheet: "", range: "$C$2:$C$5" });
  });
});
