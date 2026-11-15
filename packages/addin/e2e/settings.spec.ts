import { chunks, expect, test } from "./fixtures/test";

/** Indstillinger, der påvirker, hvordan modellen kaldes. */

test("model pr. opgave (#13): valgt model bruges kun til den opgave", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Lang og indviklet tekst.", models: ["qwen3:4b", "qwen3:32b"] });
  tp.model.reply = () => ({ stream: chunks("Kort tekst.") });

  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  await tp.page.getByRole("button", { name: /Model pr\. opgave/ }).click();
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
