import { describe, expect, it } from "vitest";
import { launchInfo, secondInstanceDialog, whereToFind } from "../instance.js";

describe("MIAI startet, mens en udgave allerede kører", () => {
  it("samme udgave: sig, hvor MIAI er – på den platforms sprog", () => {
    const d = secondInstanceDialog("0.1.9", launchInfo("0.1.9", "/Applications/MIAI.app/Contents/MacOS/MIAI"), "darwin");
    expect(d).toMatchObject({ kind: "running", message: "MIAI kører allerede", buttons: ["OK"] });
    expect(d.detail).toContain("menulinjen");
    expect(secondInstanceDialog("0.1.9", launchInfo("0.1.9", "C:\\MIAI.exe"), "win32").detail).toContain("proceslinjen");
  });

  it("anden udgave: tilbyd at skifte og husk, hvilken der skal startes", () => {
    const d = secondInstanceDialog("0.1.8", launchInfo("0.1.9", "/Applications/MIAI.app/Contents/MacOS/MIAI"), "darwin");
    expect(d).toEqual({
      kind: "switch",
      message: "Skift til MIAI 0.1.9?",
      detail: "MIAI 0.1.8 kører allerede, så 0.1.9 kunne ikke starte. Skift lukker 0.1.8 og starter 0.1.9.",
      buttons: ["Skift til 0.1.9", "Behold 0.1.8"],
      execPath: "/Applications/MIAI.app/Contents/MacOS/MIAI",
    });
  });

  it("en ældre udgave sender intet med: behandles som samme udgave", () => {
    for (const other of [undefined, null, {}, { version: "0.1.9" }, { execPath: "/x" }, "tekst"])
      expect(secondInstanceDialog("0.1.9", other, "darwin").kind, JSON.stringify(other)).toBe("running");
  });

  it("stedet, MIAI findes", () => {
    expect(whereToFind("darwin")).toBe("i menulinjen øverst til højre");
    expect(whereToFind("win32")).toBe("i proceslinjen ved uret");
  });
});
