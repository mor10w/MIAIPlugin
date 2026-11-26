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

test("agenten mangler tal fra nettet: chatten slår dem op, og agenten laver planen ud fra dem", async ({ open }) => {
  const tp = await open({
    host: "Excel",
    mode: "bridge",
    apiSets: { ExcelApi: "1.10" },
    excel: {
      sheet: "Aktier",
      values: [
        ["Aktie", "Kurs"],
        ["NOVO-B", 274.3],
      ],
    },
    settings: { webSearch: true, search: { provider: "searxng", url: searxUrl } },
  });
  const table = "| Dato | NOVO-B |\n|---|---|\n| 2026-11-24 | 274.65 |\n| 2026-11-25 | 274.30 |\n\nKilde: ollama.com";
  tp.model.reply = (b) => {
    const agent = String(b.messages[0].content).includes("You are in agent mode");
    const last = String(b.messages.at(-1)!.content);
    if (agent && !last.includes("<material>"))
      return { json: JSON.stringify({ summary: "Jeg har ikke kurserne.", actions: [], needs: "lukkekurser for NOVO-B de seneste dage" }) };
    if (agent)
      return {
        json: JSON.stringify({
          summary: "Skriver kurserne ved siden af tabellen.",
          actions: [
            {
              action: "setValues",
              address: "D1",
              values: [
                ["Dato", "NOVO-B"],
                ["2026-11-24", 274.65],
                ["2026-11-25", 274.3],
              ],
            },
          ],
        }),
      };
    return b.messages.some((m) => m.role === "tool") ? { stream: [table] } : { toolCall: { name: "web_search", args: { query: "NOVO-B lukkekurser" } } };
  };

  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Tilføj de seneste dages lukkekurser til aktien i tabellen");
  await tp.page.keyboard.press("Enter");

  // Agentens afvisning vises ikke som en blindgyde: chatten søger, svaret står i samtalen, og planen følger.
  await expect(tp.page.getByText(/Agenten mangler oplysninger .* lukkekurser for NOVO-B de seneste dage/)).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Udfør 1 handling" })).toBeVisible();
  await expect(tp.page.getByText("Skriver kurserne ved siden af tabellen.")).toBeVisible();
  await expect(tp.page.getByText("Jeg har ikke kurserne.")).toHaveCount(0);
  expect(queries).toContain("NOVO-B lukkekurser");

  // Chatten fik at vide, hvad agenten mangler; agenten fik chattens tabel som materiale.
  const lookup = tp.model.requests.find((r) => String(r.messages[0].content).includes("lacks this data: lukkekurser for NOVO-B"));
  expect(lookup).toBeDefined();
  expect(String(tp.model.lastRequest!.messages.at(-1)!.content)).toContain("| 2026-11-25 | 274.30 |");
  // Kun ét opslag: tre kald til agent-opgaven ville være en løkke.
  expect(tp.model.requests.filter((r) => String(r.messages[0].content).includes("You are in agent mode"))).toHaveLength(2);

  await tp.page.getByRole("button", { name: "Udfør 1 handling" }).click();
  await expect(tp.page.getByText("1 af 1 handlinger udført.")).toBeVisible();
});

test("uden web-søgning vises agentens forklaring som før", async ({ open }) => {
  const tp = await open({ host: "Excel", mode: "bridge", apiSets: { ExcelApi: "1.10" }, excel: { sheet: "Aktier", values: [["Aktie"], ["NOVO-B"]] } });
  tp.model.reply = () => ({ json: JSON.stringify({ summary: "Jeg har ikke kurserne.", actions: [], needs: "lukkekurser for NOVO-B" }) });
  await tp.tab("Chat");
  await tp.page.getByPlaceholder("Spørg om dit dokument…").fill("Tilføj de seneste dages lukkekurser til aktien i tabellen");
  await tp.page.keyboard.press("Enter");
  await expect(tp.page.getByText("Jeg har ikke kurserne.")).toBeVisible();
  expect(tp.model.requests).toHaveLength(1);
});
