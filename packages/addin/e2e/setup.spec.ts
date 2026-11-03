import { portIsFree } from "./fixtures/servers";
import { expect, test } from "./fixtures/test";

/**
 * Førstegangsguide (#4). Mock-modellen lytter på værktøjernes kendte porte (Ollama 11434,
 * LM Studio 1234), så testene kører serielt og springes over, hvis portene er optaget lokalt.
 */
test.describe.configure({ mode: "serial" });

const DEAD = { baseUrl: "http://127.0.0.1:9/v1" }; // ingen svarer på port 9

test.beforeEach(async () => {
  test.skip(!(await portIsFree(11434)) || !(await portIsFree(1234)), "Port 11434 eller 1234 er optaget (kører Ollama/LM Studio lokalt?)");
});

test("uden forbindelse: Kom i gang → guiden finder Ollama → Brug", async ({ open }) => {
  const tp = await open({ host: "Word", modelPort: 11434, endpoint: DEAD, models: ["qwen3:4b", "gemma3:12b"] });
  await tp.page.getByRole("button", { name: "Kom i gang" }).click();

  await expect(tp.page.getByText("Klar – 2 modeller")).toBeVisible();
  await tp.page.getByRole("button", { name: "Brug", exact: true }).click();
  await expect(tp.page.getByText("Forbundet til Ollama – 2 modeller fundet.")).toBeVisible();
  // Statusmærket opdateres med det samme, og endpointen er gemt som aktiv.
  await expect(tp.page.getByText("Forbundet", { exact: true })).toBeVisible();
  const saved = await tp.page.evaluate(() => JSON.parse(localStorage.getItem("miai.settings.v1")!));
  expect(saved.endpoints.find((e: { id: string }) => e.id === saved.activeEndpoint).baseUrl).toBe("http://127.0.0.1:11434/v1");
});

test("værktøj uden CORS: præcis opsætning med add-in'ets origin", async ({ open }) => {
  const tp = await open({ host: "Word", modelPort: 1234, cors: false, endpoint: DEAD });
  await tp.tab("Indstillinger");
  const guide = tp.page.locator(".fui-Card").filter({ hasText: "Find lokale værktøjer" });
  await expect(guide.getByText("Kører – mangler CORS")).toBeVisible();
  await expect(guide.getByText(/Enable CORS/)).toBeVisible();
  const origin = new URL(tp.page.url()).origin;
  await expect(guide.getByText(`tillader ikke add-in'et (${origin})`, { exact: false })).toBeVisible();
});

test("Ollama uden CORS: kopiérbar kommando med den præcise origin (ikke *)", async ({ open }) => {
  const tp = await open({ host: "Word", modelPort: 11434, cors: false, endpoint: DEAD });
  await tp.tab("Indstillinger");
  const origin = new URL(tp.page.url()).origin;
  await expect(tp.page.getByText(`launchctl setenv OLLAMA_ORIGINS "${origin}"`)).toBeVisible();
  await expect(tp.page.getByText('OLLAMA_ORIGINS "*"')).toHaveCount(0);
});

test("Gem validerer endpointen først – og kan gemme alligevel", async ({ open }) => {
  const tp = await open({ host: "Word" }); // virkende mock på tilfældig port
  await tp.tab("Indstillinger");
  const url = tp.page.getByLabel("Basis-URL (inkl. /v1)");
  await url.fill("http://127.0.0.1:9/v1");
  await tp.page.getByRole("button", { name: "Gem", exact: true }).click();
  await expect(tp.page.getByText(/Endpointen svarer ikke/)).toBeVisible();
  // Intet gemt endnu.
  expect(await tp.page.evaluate(() => JSON.parse(localStorage.getItem("miai.settings.v1")!).endpoints[0].baseUrl)).not.toContain(":9/");
  await tp.page.getByRole("button", { name: "Gem alligevel" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();
  expect(await tp.page.evaluate(() => JSON.parse(localStorage.getItem("miai.settings.v1")!).endpoints[0].baseUrl)).toBe("http://127.0.0.1:9/v1");
});

test("bridge-tilstand: guiden finder værktøjet uden CORS (bridgen tester selv)", async ({ open }) => {
  const tp = await open({ host: "Word", mode: "bridge", modelPort: 11434, cors: false, endpoint: DEAD });
  await tp.tab("Indstillinger");
  await expect(tp.page.getByText("Klar – 1 modeller")).toBeVisible();
});
