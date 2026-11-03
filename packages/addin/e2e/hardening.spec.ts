import { chunks, expect, test } from "./fixtures/test";

/** Hærdning (#3): lange dokumenter, kontekstgrænser og forståelige fejl. */

const page = (n: number) => `Side ${n}. ${"Projektet følger planen, og budgettet holder. ".repeat(60)}`;
const document = (pages: number) => Array.from({ length: pages }, (_, i) => page(i + 1)).join("\n\n");

test("opsummering af et dokument på 50+ sider deles op og samles", async ({ open }) => {
  const doc = document(60); // ca. 170.000 tegn
  const tp = await open({ host: "Word", selection: doc, endpoint: { contextTokens: 8_192 } });
  tp.model.reply = (b) => (b.messages[0].content.includes("one part of a longer document") ? { stream: ["Del-resumé."] } : { stream: chunks("Samlet resumé af hele dokumentet.") });

  await tp.run("Opsummér");
  await expect(tp.page.getByText("Samlet resumé af hele dokumentet.")).toBeVisible({ timeout: 20_000 });
  // Hele dokumentet blev sendt (ingen afkortning), fordelt på flere dele + én samlet kørsel.
  const parts = tp.model.requests.filter((r) => r.messages[0].content.includes("one part of a longer document"));
  expect(parts.length).toBeGreaterThan(5);
  expect(parts.reduce((n, r) => n + (r.messages.at(-1)?.content.length ?? 0), 0)).toBeGreaterThan(doc.length * 0.99);
  await expect(tp.page.getByText(/Kun de første/)).toHaveCount(0);
});

test("afkortet dokument meldes til brugeren", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "x".repeat(50_000) });
  tp.model.reply = () => ({ stream: ["Kortere."] });
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Kun de første 24.000 af 50.000 tegn blev sendt til modellen.")).toBeVisible();
});

test("advarsel, når modellen ikke fik hele teksten (usage viser færre tokens)", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Lang tekst. ".repeat(1_500) });
  tp.model.reply = () => ({ stream: ["Resumé."], promptTokens: 2_048 }); // som Ollama med num_ctx 2048
  await tp.run("Forenkl");
  await expect(tp.page.getByText(/Modellen fik ikke hele teksten med/)).toBeVisible();
  expect(tp.model.lastRequest?.stream_options).toEqual({ include_usage: true });
});

test("advarsel, når svaret ramte grænsen for svarlængde", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Tekst." });
  tp.model.reply = () => ({ stream: ["Et svar, der stopper midt i"], finishReason: "length" });
  await tp.run("Forenkl");
  await expect(tp.page.getByText(/Svaret blev afkortet/)).toBeVisible();
});

test("401 fra endpointen giver en forståelig fejl med serverens detalje", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "Tekst." });
  tp.model.reply = () => ({ status: 401, error: "invalid api key" });
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Endpointen afviste adgang – tjek API-nøglen.")).toBeVisible();
  await expect(tp.page.locator(".fui-MessageBarBody")).toContainText("invalid api key");
});

test("kendt kontekst: for lang tekst afvises med det samme", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "x".repeat(20_000), endpoint: { contextTokens: 2_048 } });
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Teksten er for lang til modellen.")).toBeVisible();
  expect(tp.model.requests).toHaveLength(0); // intet spildt kald
});
