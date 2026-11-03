import type { OfficeCall } from "./fixtures/office";
import { expect, test } from "./fixtures/test";

/** "Lav diagram" (ADR-0009) i PowerPoint (native figurer) og Word (billede). */

// Feedback-løkke (Rapport → Indsamling): lagde tidligere hele diagrammet i én søjle uden pilehoveder.
const DIAGRAM = {
  title: "Proces",
  nodes: [
    { id: "n1", label: "Indsamling af data fra alle afdelinger og leverandører" },
    { id: "n2", label: "Analyse" },
    { id: "n3", label: "Kvalitetssikring" },
    { id: "n4", label: "Rapport til ledelsen" },
  ],
  edges: [
    { from: "n1", to: "n2" },
    { from: "n1", to: "n3" },
    { from: "n2", to: "n4" },
    { from: "n3", to: "n4" },
    { from: "n4", to: "n1" },
  ],
};

const count = (calls: OfficeCall[], kind: string, type?: string) => calls.filter((c) => c.kind === kind && (!type || c.type === type)).length;

for (const pp18 of [false, true]) {
  test(`PowerPoint: native figurer med vinkelrette linjer og pile${pp18 ? " – grupperet (PowerPointApi 1.8)" : ""}`, async ({ open }) => {
    const tp = await open({ host: "PowerPoint", apiSets: { PowerPointApi: pp18 ? "1.8" : "1.4" }, selection: "Proces: indsamling, analyse, QA, rapport" });
    tp.model.reply = () => ({ json: JSON.stringify(DIAGRAM) });
    await tp.run("Lav diagram");

    // Forhåndsvisningen er gyldig SVG med ombrudte labels som <tspan>.
    const svg = tp.page.locator("svg[viewBox]").first();
    await expect(svg).toBeVisible();
    expect(await svg.locator("tspan").count()).toBeGreaterThan(DIAGRAM.nodes.length);
    expect(await svg.textContent()).not.toContain("tspan");

    await tp.page.getByRole("button", { name: "Indsæt diagram" }).click();
    await expect(tp.page.getByText(/4 figurer og 5 forbindelser/)).toBeVisible();

    const calls = await tp.calls();
    expect(count(calls, "shape", "RoundRectangle")).toBe(DIAGRAM.nodes.length);
    expect(count(calls, "shape", "Chevron")).toBe(DIAGRAM.edges.length); // én pil pr. kant
    // PowerPoint kan ikke vende linjer: hvert segment skal være vandret eller lodret.
    const lines = calls.filter((c) => c.kind === "line") as unknown as { width: number; height: number }[];
    expect(lines.length).toBeGreaterThan(0);
    expect(lines.filter((l) => l.width > 0 && l.height > 0)).toEqual([]);
    expect(count(calls, "group")).toBe(pp18 ? 1 : 0);
  });
}

test("Word: billede i naturlig størrelse med alt-tekst", async ({ open }) => {
  const tp = await open({ host: "Word", apiSets: { WordApi: "1.4" }, selection: "Proces: indsamling, analyse, QA, rapport" });
  tp.model.reply = () => ({ json: JSON.stringify(DIAGRAM) });
  await tp.run("Lav diagram");
  await tp.page.getByRole("button", { name: "Indsæt diagram" }).click();
  await expect(tp.page.getByText("Figuren er indsat efter markeringen.")).toBeVisible();

  const pic = (await tp.find<{ bytes: number; width: number; height: number; altTextTitle: string; altTextDescription: string }>("picture"))!;
  expect(pic.bytes).toBeGreaterThan(1_000);
  expect(pic.width).toBeLessThanOrEqual(450);
  expect(pic.height).toBeGreaterThan(0);
  expect(pic.altTextTitle).toBe("Proces");
  expect(pic.altTextDescription).toContain("Analyse");
});
