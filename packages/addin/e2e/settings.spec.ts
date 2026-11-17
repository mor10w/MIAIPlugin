import type { Locator } from "@playwright/test";
import { chunks, expect, test } from "./fixtures/test";

/** Indstillinger, der påvirker, hvordan modellen kaldes. */

test("model pr. opgave (#13): valgt model bruges kun til den opgave", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Lang og indviklet tekst.", models: ["qwen3:4b", "qwen3:32b"] });
  tp.model.reply = () => ({ stream: chunks("Kort tekst.") });

  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  await tp.page.getByRole("button", { name: /Undtagelser for enkelte opgaver/ }).click();
  await tp.page.getByRole("combobox", { name: "Forenkl" }).click();
  await tp.page.getByRole("option", { name: "qwen3:32b" }).click();
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();

  await tp.tab("Handlinger");
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Kort tekst.")).toBeVisible();
  expect(tp.model.lastRequest?.model).toBe("qwen3:32b");
  await expect(tp.page.getByText("qwen3:32b")).toBeVisible(); // modellen vises ved resultatet

  await tp.run("Parafrasér");
  await expect(tp.page.getByText("qwen3:4b")).toBeVisible();
  expect(tp.model.lastRequest?.model).toBe("qwen3:4b"); // standard: første model fra endpointen
});

test("model pr. opgavetype: én model til Ret og omskriv gælder alle dens opgaver; en undtagelse har forrang", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Lang og indviklet tekst.", models: ["qwen3:4b", "qwen3:32b", "gemma3:12b"] });
  tp.model.reply = () => ({ stream: chunks("Kort tekst.") });

  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  // Typerne vises med det samme – uden at folde noget ud.
  for (const name of ["Ret og omskriv", "Skriv nyt", "Struktur, formler og data", "Agent", "Chat"])
    await expect(tp.page.getByRole("combobox", { name })).toBeVisible();
  await tp.page.getByRole("combobox", { name: "Ret og omskriv" }).click();
  await tp.page.getByRole("option", { name: "qwen3:32b" }).click();
  // Undtagelse: Oversæt skal bruge en tredje model.
  await tp.page.getByRole("button", { name: /Undtagelser for enkelte opgaver/ }).click();
  await tp.page.getByRole("combobox", { name: "Oversæt" }).click();
  await tp.page.getByRole("option", { name: "gemma3:12b" }).click();
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();

  await tp.tab("Handlinger");
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Kort tekst.")).toBeVisible();
  expect(tp.model.lastRequest?.model).toBe("qwen3:32b"); // typens model
  await tp.run("Oversæt");
  await expect(tp.page.getByText("gemma3:12b")).toBeVisible();
  expect(tp.model.lastRequest?.model).toBe("gemma3:12b"); // undtagelsen

  tp.model.reply = () => ({ stream: chunks("Svar i chatten.") });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad står der?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Svar i chatten.")).toBeVisible();
  expect(tp.model.lastRequest?.model).toBe("qwen3:4b"); // Chat har intet valg: standard (første model)
});

test("model pr. opgave: advarsel, når modellen ikke findes på endpointen", async ({ open }) => {
  const tp = await open({ host: "Word", settings: { taskModels: { simplify: "fjernet-model" } } });
  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  // Listen over model pr. opgave er åben, når der er valg.
  await expect(tp.page.getByText(/Modellen findes ikke på den aktive endpoint/)).toBeVisible();
});

test("timeout kan indstilles pr. endpoint (10–1800 s) og gemmes", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.page.getByRole("button", { name: "Avanceret" }).click();
  const field = tp.page.getByRole("spinbutton", { name: "Timeout" });
  // Standard: 5 minutter.
  await expect(field).toHaveValue("300");
  // Under minimum rettes til 10 s, når feltet forlades.
  await field.fill("5");
  await field.press("Tab");
  await expect(field).toHaveValue("10");
  await field.fill("900");
  await field.press("Enter");
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();
  const saved = await tp.page.evaluate(() => JSON.parse(localStorage.getItem("miai.settings.v1") ?? "{}"));
  expect(saved.endpoints[0].timeoutMs).toBe(900_000);
});

test("tænkning og indsats: valget sendes til modellen, og afviser serveren det, vises en besked", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Lang og indviklet tekst." });
  const FIELDS = ["reasoning_effort", "think", "chat_template_kwargs"];
  const choose = async (label: string, option: string) => {
    await tp.page.getByRole("combobox", { name: label }).click();
    await tp.page.getByRole("option", { name: option, exact: true }).click();
  };
  const save = async () => {
    await tp.page.getByRole("button", { name: "Gem" }).click();
    await expect(tp.page.getByText("Gemt.")).toBeVisible();
  };

  // Serverens standard: intet af det sendes.
  tp.model.reply = () => ({ stream: chunks("Første svar.") });
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Første svar.")).toBeVisible();
  for (const f of FIELDS) expect(tp.model.lastRequest).not.toHaveProperty(f);

  // Fra: sendes i alle tre dialekter, og indsatsen kan ikke vælges.
  await tp.tab("Indstillinger");
  await tp.section("Svar");
  await choose("Tænkning", "Fra");
  await expect(tp.page.getByRole("combobox", { name: "Indsats" })).toBeDisabled();
  await save();
  tp.model.reply = () => ({ stream: chunks("Andet svar.") });
  await tp.tab("Handlinger");
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Andet svar.")).toBeVisible();
  expect(tp.model.lastRequest).toMatchObject({ reasoning_effort: "none", think: false, chat_template_kwargs: { enable_thinking: false } });

  // Til med høj indsats.
  await tp.tab("Indstillinger");
  await tp.section("Svar");
  await choose("Tænkning", "Til");
  await choose("Indsats", "Høj");
  await save();
  tp.model.reply = () => ({ stream: chunks("Tredje svar.") });
  await tp.tab("Handlinger");
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Tredje svar.")).toBeVisible();
  expect(tp.model.lastRequest).toMatchObject({ reasoning_effort: "high", chat_template_kwargs: { enable_thinking: true } });
  expect(tp.model.lastRequest).not.toHaveProperty("think");
  await expect(tp.page.getByText(/understøtter ikke dit valg af tænkning/)).toHaveCount(0);

  // En streng server, der afviser felterne: svaret laves uden dem, og brugeren får besked.
  tp.model.reply = (body) =>
    FIELDS.some((f) => f in body) ? { status: 400, error: "Unrecognized request argument supplied: reasoning_effort" } : { stream: chunks("Fjerde svar.") };
  await tp.run("Parafrasér");
  await expect(tp.page.getByText("Fjerde svar.")).toBeVisible();
  for (const f of FIELDS) expect(tp.model.lastRequest).not.toHaveProperty(f);
  await expect(tp.page.getByText(/understøtter ikke dit valg af tænkning/)).toBeVisible();
});

/** Felter side om side (twoCol): samme højde og inden for kortet, også uden hint og med en Dropdown. */
async function expectSideBySide(left: Locator, right: Locator) {
  await expect(right).toBeVisible();
  // Accordionen glider op: mål igen, til felterne er på plads (med fejlen bliver de det aldrig).
  await expect(async () => {
    const [l, r] = [await left.boundingBox(), await right.boundingBox()];
    const card = await right.locator("xpath=ancestor::*[contains(@class,'fui-AccordionItem')][1]").boundingBox();
    expect(l && r && card).toBeTruthy();
    expect(Math.abs(l!.y - r!.y)).toBeLessThan(1);
    expect(Math.abs(l!.height - r!.height)).toBeLessThan(1);
    expect(l!.x + l!.width).toBeLessThanOrEqual(r!.x);
    expect(r!.x + r!.width).toBeLessThanOrEqual(card!.x + card!.width);
  }).toPass({ timeout: 5_000 });
}

test("Svar: Temperatur og Maks. tokens flugter og holder sig inden for kortet", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.page.getByRole("button", { name: "Forbindelse", exact: true }).click();
  await tp.section("Svar");
  const field = (name: string) => tp.page.locator(".fui-Input", { has: tp.page.getByRole("spinbutton", { name }) });
  await expectSideBySide(field("Temperatur"), field("Maks. tokens i svar"));
});

test("Søgetjeneste og SearXNG-adresse flugter og holder sig inden for kortet", async ({ open }) => {
  const tp = await open({ host: "Word", mode: "bridge", settings: { webSearch: true, search: { provider: "searxng", url: "http://127.0.0.1:1" } } });
  await tp.tab("Indstillinger");
  await tp.page.getByRole("button", { name: "Forbindelse", exact: true }).click();
  await tp.section("Privatliv og politik");
  const searxng = tp.page.locator(".fui-Input", { has: tp.page.getByRole("textbox", { name: "SearXNG-adresse" }) });
  await searxng.scrollIntoViewIfNeeded();
  const provider = tp.page.locator(".fui-Dropdown", { has: tp.page.getByRole("combobox", { name: "Søgetjeneste" }) });
  await expectSideBySide(provider, searxng);
});
