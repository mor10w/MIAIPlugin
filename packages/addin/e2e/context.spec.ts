import { chunks, expect, test } from "./fixtures/test";

/** Modellens kontekst (#73): aflæses fra Ollama og LM Studio, vises under Indstillinger og advarer i chat og agent. */

const ollama = (loaded?: number) => ({
  ollama: {
    show: {
      parameters: "",
      model_info: { "general.architecture": "qwen3", "qwen3.context_length": 40_960 },
      capabilities: ["completion", "tools", "thinking"],
    },
    ps: { models: loaded ? [{ name: "qwen3:4b", model: "qwen3:4b", context_length: loaded }] : [] },
  },
});

for (const mode of ["direct", "bridge"] as const) {
  test(`Ollama med 4.096 tokens: konteksten og en rettelse vises under Indstillinger (${mode})`, async ({ open }) => {
    const tp = await open({ host: "Word", mode, native: ollama(4_096) });
    await tp.tab("Indstillinger");
    await tp.section("Modeller");
    const box = tp.page.getByTestId("model-context");
    await expect(box).toContainText("4.096 tokens – som modellen er indlæst med i Ollama.");
    await expect(box).toContainText("Modellen kan højst 40.960.");
    await expect(box).toContainText("Kan: værktøjer, tænkning.");
    await expect(box).toContainText("For lidt til agent og chat, som har brug for mindst 16.384.");
    await expect(box).toContainText("OLLAMA_CONTEXT_LENGTH");
  });
}

test("Ollama, hvor modellen ikke er indlæst endnu: forklaring og Tjek igen", async ({ open }) => {
  const tp = await open({ host: "Word", native: ollama() });
  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  const box = tp.page.getByTestId("model-context");
  await expect(box).toContainText("Kendes først, når modellen er indlæst");
  await expect(box).toContainText("4.096 tokens på maskiner med under 24 GB grafikhukommelse");
  await box.getByRole("button", { name: "Tjek igen" }).click();
  await expect(box).toContainText("Kendes først, når modellen er indlæst");
});

test("LM Studio: den indlæste kontekst, vision og værktøjer; ingen advarsel ved stor kontekst", async ({ open }) => {
  const tp = await open({
    host: "Word",
    models: ["qwen/qwen3-8b"],
    native: {
      lmstudio: {
        models: [
          {
            type: "llm",
            key: "qwen/qwen3-8b",
            max_context_length: 131_072,
            loaded_instances: [{ id: "qwen/qwen3-8b", config: { context_length: 32_768 } }],
            capabilities: { vision: true, trained_for_tool_use: true },
          },
        ],
      },
    },
  });
  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  const box = tp.page.getByTestId("model-context");
  await expect(box).toContainText("32.768 tokens – som modellen er indlæst med i LM Studio. Modellen kan højst 131.072.");
  await expect(box).toContainText("Kan: billeder, værktøjer.");
  await expect(box).not.toContainText("For lidt");
});

test("et værktøj uden kontekstoplysning: forklaring om feltet Kontekstlængde", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.section("Modeller");
  await expect(tp.page.getByTestId("model-context")).toContainText("Ukendt. Værktøjet oplyser ikke konteksten");
});

test("chat på en model med 4.096 tokens: advarsel med rettelsen til Ollama", async ({ open }) => {
  const tp = await open({ host: "Word", native: ollama(4_096) });
  tp.model.reply = () => ({ stream: chunks("Kort svar.") });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad handler det om?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Kort svar.")).toBeVisible();
  await expect(tp.page.getByText(/Modellen kører med kun 4\.096 tokens kontekst\. Agent og chat har brug for mindst 16\.384/)).toBeVisible();
  await expect(tp.page.getByText(/OLLAMA_CONTEXT_LENGTH/)).toBeVisible();
});
