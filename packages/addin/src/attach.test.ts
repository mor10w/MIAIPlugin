import { describe, expect, it } from "vitest";
import { AttachmentError, clipText, fitImage, toAttachment } from "./attach";

const file = (name: string, type: string, content = "indhold") => new File([content], name, { type });

describe("vedhæftninger i chatten (#27)", () => {
  it("billeder skaleres, så længste side højst er 1568 px", () => {
    expect(fitImage(3136, 1568)).toEqual({ width: 1568, height: 784 });
    expect(fitImage(800, 600)).toEqual({ width: 800, height: 600 });
    expect(fitImage(100, 5000)).toEqual({ width: 31, height: 1568 });
  });

  it("tekstfiler læses som tekst og afkortes med en note", async () => {
    expect(await toAttachment(file("noter.txt", "text/plain", "Budget 2026"))).toEqual({ kind: "text", name: "noter.txt", text: "Budget 2026" });
    expect(await toAttachment(file("data.csv", "", "a;b\n1;2"))).toMatchObject({ kind: "text", text: "a;b\n1;2" });
    const long = clipText("x".repeat(70_000));
    expect(long).toHaveLength(60_000);
    expect(long.endsWith("[… the rest of the file was left out]")).toBe(true);
  });

  it("PDF og Office-filer kræver bridgens udtræk; ukendte typer afvises", async () => {
    await expect(toAttachment(file("rapport.pdf", "application/pdf"))).rejects.toMatchObject({ reason: "needs-bridge", file: "rapport.pdf" });
    expect(await toAttachment(file("rapport.pdf", "application/pdf"), async () => "Udtrukket tekst")).toEqual({
      kind: "text",
      name: "rapport.pdf",
      text: "Udtrukket tekst",
    });
    await expect(toAttachment(file("program.exe", "application/octet-stream"))).rejects.toBeInstanceOf(AttachmentError);
    await expect(toAttachment(file("program.exe", "application/octet-stream"))).rejects.toMatchObject({ reason: "unsupported" });
  });
});
