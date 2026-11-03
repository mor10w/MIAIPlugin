import { chunks, expect, test } from "./fixtures/test";

/** Forbindelse og grundflow i begge tilstande (ADR-0007). */

test.describe("direkte tilstand", () => {
  test("forbinder, streamer, filtrerer <think> og indsætter i Word", async ({ open }) => {
    const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, selection: "Den foreliggende tekst er af en overordentlig kompliceret beskaffenhed." });
    tp.model.reply = () => ({ stream: ["<thi", "nk>overvejer…</think>\n\n", ...chunks("Dette er en **enkel** tekst.")] });

    await expect(tp.page.getByText("Forbundet")).toBeVisible();
    await tp.run("Forenkl");
    // Ren tekst-opgave: Markdown-fremhævning fjernes, tankegangen vises ikke.
    await expect(tp.page.getByText("Dette er en enkel tekst.")).toBeVisible();
    await expect(tp.page.getByText("overvejer")).toHaveCount(0);

    // Regression: en Fluent-portal dækkede tidligere hele panelet, så knapper ikke kunne klikkes.
    await tp.page.getByRole("button", { name: "Erstat" }).click();
    await expect(tp.page.getByText("Indsat.")).toBeVisible();
    expect(await tp.calls()).toContainEqual(expect.objectContaining({ kind: "insertText", text: "Dette er en enkel tekst.", loc: "Replace" }));

    // Dokumentindholdet sendes som data i <document>-tags.
    expect(tp.model.lastRequest?.messages.at(-1)?.content).toContain("<document>");
  });

  test("uden CORS: tydelig fejl med vejledning", async ({ open }) => {
    const tp = await open({ host: "Word", selection: "Tekst", cors: false });
    await expect(tp.page.getByText("Ingen forbindelse til modellen")).toBeVisible();
    await tp.run("Forenkl");
    await expect(tp.page.locator(".fui-MessageBarBody")).toContainText(/tillader CORS fra https:\/\/localhost:\d+/);
  });

  test("grammatik indsættes som sporede ændringer", async ({ open }) => {
    const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, selection: "Jeg har en hun og en kat." });
    tp.model.reply = () => ({ json: JSON.stringify({ edits: [{ original: "hun", replacement: "hund", explanation: "Stavefejl" }] }) });
    await tp.run("Ret stavning og grammatik");
    await expect(tp.page.getByText("Stavefejl")).toBeVisible();
    await tp.page.getByRole("button", { name: "Indsæt som sporede ændringer" }).click();
    await expect(tp.page.getByText("1 af 1 rettelser anvendt som sporede ændringer.")).toBeVisible();
    expect(await tp.calls()).toContainEqual(expect.objectContaining({ kind: "replace", from: "hun", to: "hund" }));
    expect(tp.model.lastRequest?.response_format?.type).toBe("json_schema");
  });
});

test.describe("bridge-tilstand", () => {
  test("detekteres automatisk og virker uden CORS", async ({ open }) => {
    const tp = await open({ host: "Word", mode: "bridge", selection: "Lang og indviklet tekst.", cors: false });
    tp.model.reply = () => ({ stream: chunks("Kort tekst.") });
    await expect(tp.page.getByText("Forbundet")).toBeVisible();
    await tp.run("Forenkl");
    await tp.page.getByRole("button", { name: "Erstat" }).click();
    expect(await tp.calls()).toContainEqual(expect.objectContaining({ kind: "insertText", text: "Kort tekst." }));
  });
});
