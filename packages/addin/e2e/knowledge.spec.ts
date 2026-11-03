import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import AxeBuilder from "@axe-core/playwright";
import { expect, test } from "./fixtures/test";

/** Lokal RAG (F15): mapper indekseres i bridgen, chatten svarer med kilder, og kilden kan åbnes. */

test("bridge: indeksér en mappe, få svar med kilde, se uddraget og åbn filen", async ({ open }) => {
  const docs = mkdtempSync(join(tmpdir(), "miai-e2e-viden-"));
  mkdirSync(join(docs, "Politikker"));
  writeFileSync(join(docs, "Politikker", "Rejsepolitik.md"), "# Rejsepolitik\n\nFly bruges kun ved rejser over 500 km. Ellers tog eller bil.");
  writeFileSync(join(docs, "Kantine.md"), "# Kantine\n\nKantinen serverer frokost mellem 11.30 og 13.00.");

  const tp = await open({ host: "Word", mode: "bridge", models: ["qwen3:4b", "bge-m3"] });
  await tp.tab("Indstillinger");
  await tp.section("Viden fra mine mapper");
  await tp.page.getByRole("switch", { name: "Søg i mine mapper" }).check();
  await tp.page.getByRole("textbox", { name: "Tilføj mappe" }).fill(docs);
  await tp.page.getByRole("button", { name: "Tilføj mappe" }).click();
  await expect(tp.page.getByText(docs, { exact: true })).toBeVisible();
  await tp.page.getByRole("combobox", { name: "Embedding-model" }).fill("bge-m3");
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();
  await expect(tp.page.getByText("2 filer · 2 uddrag")).toBeVisible();
  // Uddragene sendes med filnavnet foran til embedding-modellen.
  expect(tp.model.embeddingInputs).toContainEqual(expect.stringContaining("Rejsepolitik.md\n# Rejsepolitik"));

  tp.model.reply = () => ({ stream: ["Fly bruges kun ved rejser over 500 km [1]."] });
  await tp.tab("Chat");
  await expect(tp.page.getByRole("switch", { name: "Søg i mine mapper" })).toBeChecked();
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvornår må jeg tage fly på rejser?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Fly bruges kun ved rejser over 500 km [1].")).toBeVisible();

  // Modellen fik uddraget som data med nummer.
  const sys = tp.model.lastRequest!.messages[0].content;
  expect(sys).toMatch(/<source n="1" title="Rejsepolitik.md"[\s\S]*Fly bruges kun ved rejser over 500 km/);

  // Kilden vises under svaret; uddraget og stien på klik; filen kan åbnes.
  await expect(tp.page.getByText("Kilder fra dine mapper")).toBeVisible();
  const source = tp.page.getByRole("button", { name: /^\[1\]\s*Rejsepolitik\.md$/ });
  await source.click();
  await expect(source).toHaveAttribute("aria-expanded", "true");
  await expect(tp.page.getByText(join(docs, "Politikker", "Rejsepolitik.md"))).toBeVisible();
  // Kildelisten er tilgængelig (samme regler som #17).
  const axe = await new AxeBuilder({ page: tp.page }).withTags(["wcag2a", "wcag2aa", "wcag21a", "wcag21aa"]).exclude("[data-tabster-dummy]").analyze();
  expect(axe.violations.filter((v) => v.impact === "serious" || v.impact === "critical").map((v) => v.id)).toEqual([]);
  await tp.page.getByRole("button", { name: "Åbn fil" }).first().click();
  await expect.poll(() => tp.opened).toEqual([join(docs, "Politikker", "Rejsepolitik.md")]);

  // Uden tilvalget søges der ikke.
  await tp.page.getByRole("switch", { name: "Søg i mine mapper" }).uncheck();
  tp.model.reply = () => ({ stream: ["Hej."] });
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hej");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Hej.", { exact: true })).toBeVisible();
  expect(tp.model.lastRequest!.messages[0].content).not.toContain("<sources>");
});

test("direkte tilstand: søgning i mapper kræver bridgen, og chatten viser intet tilvalg", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.section("Viden fra mine mapper");
  await expect(tp.page.getByText(/kræver MIAI Bridge eller MIAI App/)).toBeVisible();
  await tp.tab("Chat");
  await expect(tp.page.getByRole("switch", { name: "Søg i mine mapper" })).toHaveCount(0);
});
