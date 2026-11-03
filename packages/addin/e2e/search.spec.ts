import { createServer, type Server } from "node:http";
import { expect, test } from "./fixtures/test";

/** Web-søgning gennem bridgen med selvhostet SearXNG (#12) – hele kæden, uden internet. */

let searx: Server;
let searxUrl = "";
const queries: string[] = [];

test.beforeAll(async () => {
  searx = createServer((req, res) => {
    queries.push(new URL(req.url ?? "/", "http://x").searchParams.get("q") ?? "");
    res.setHeader("content-type", "application/json");
    res.end(JSON.stringify({ results: [{ title: "Ollama 0.9 udgivet", url: "https://ollama.com/blog/0.9", content: "Version 0.9 kom i oktober." }] }));
  });
  await new Promise<void>((r) => searx.listen(0, "127.0.0.1", r));
  searxUrl = `http://127.0.0.1:${(searx.address() as { port: number }).port}`;
});
test.afterAll(() => searx.close());

test("modellen søger via SearXNG; indledningen fra søgerunden fjernes fra svaret", async ({ open }) => {
  const tp = await open({ host: "Word", mode: "bridge", settings: { webSearch: true, search: { provider: "searxng", url: searxUrl } } });
  tp.model.reply = (b) =>
    b.messages.some((m) => m.role === "tool")
      ? { stream: ["Ifølge ollama.com kom version 0.9 i oktober."] }
      : { text: "Lad mig søge efter det. ", toolCall: { name: "web_search", args: { query: "ollama nyeste version" } } };

  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Hvad er nyeste Ollama-version?");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Ifølge ollama.com kom version 0.9 i oktober.")).toBeVisible();
  await expect(tp.page.getByText(/Lad mig søge/)).toHaveCount(0); // trukket tilbage (retract)

  expect(queries).toContain("ollama nyeste version"); // bridgen søgte hos SearXNG, ikke DuckDuckGo
  const toolMsg = tp.model.lastRequest!.messages.find((m) => m.role === "tool");
  expect(toolMsg?.content).toContain("https://ollama.com/blog/0.9");
  expect(tp.model.requests[0].tools).toBeDefined(); // værktøjet tilbydes, når søgning er slået til
});
