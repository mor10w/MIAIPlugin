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
