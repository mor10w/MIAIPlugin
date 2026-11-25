import { createServer, type Server } from "node:http";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { checkSettings, effectiveSearch, isUrlAllowed, Policy, Settings } from "@miai/shared";
import { BOT_CHALLENGE_TEXT, looksLikeBotChallenge, looksLikeUnparsedResults, parseDdgHtml, searchWeb } from "../src/search.js";

/** Web-søgning med selvhostet SearXNG og robust DuckDuckGo-aflæsning (#12). */

let server: Server;
let base: string;
let lastUrl = "";
let jsonEnabled = true;

beforeAll(async () => {
  server = createServer((req, res) => {
    lastUrl = req.url ?? "";
    if (!jsonEnabled) {
      res.statusCode = 403;
      return res.end("Forbidden");
    }
    res.setHeader("content-type", "application/json");
    res.end(
      JSON.stringify({
        results: [
          { title: "Ollama 0.9 udgivet", url: "https://ollama.com/blog", content: "Ny version med\n større kontekst." },
          { title: "Uden URL", content: "springes over" },
          { title: "LM Studio", url: "https://lmstudio.ai", content: "" },
        ],
      }),
    );
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  base = `http://127.0.0.1:${(server.address() as { port: number }).port}`;
});
afterAll(() => server.close());

const allowAll = () => true;

describe("SearXNG", () => {
  it("søger via JSON-API'et og formaterer resultaterne", async () => {
    jsonEnabled = true;
    const r = await searchWeb({ provider: "searxng", url: base }, "ollama version", allowAll);
    expect(lastUrl).toBe("/search?q=ollama%20version&format=json");
    expect(r.ok).toBe(true);
    expect(r.host).toBe(new URL(base).host); // audit-loggen får den faktiske søgehost
    expect(r.text).toContain("1. Ollama 0.9 udgivet\n   https://ollama.com/blog\n   Ny version med større kontekst.");
    expect(r.text).toContain("2. LM Studio");
    expect(r.text).not.toContain("Uden URL");
  });

  it("forklarer 403 (JSON-formatet er ikke slået til)", async () => {
    jsonEnabled = false;
    const r = await searchWeb({ provider: "searxng", url: base }, "x", allowAll);
    expect(r.ok).toBe(false);
    expect(r.text).toContain("search.formats");
    jsonEnabled = true;
  });

  it("respekterer politikken (allowedHosts)", async () => {
    lastUrl = "";
    const policy = Policy.parse({ allowedHosts: ["search.firma.local"] });
    const r = await searchWeb({ provider: "searxng", url: base }, "x", (u) => isUrlAllowed(u, policy));
    expect(r.ok).toBe(false);
    expect(r.text).toContain("ikke tilladt af politikken");
    expect(lastUrl).toBe(""); // intet kald
  });

  it("validering ved gem: adresse påkrævet og tilladt af politikken", () => {
    const s = (search: object) => Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: "http://127.0.0.1:1/v1" }], activeEndpoint: "x", search });
    const loopback = Policy.parse({});
    expect(checkSettings(s({ provider: "searxng" }), loopback, isUrlAllowed)).toContain("Angiv adressen");
    expect(checkSettings(s({ provider: "searxng", url: "https://search.firma.local" }), loopback, isUrlAllowed)).toContain("ikke tilladt");
    expect(
      checkSettings(
        s({ provider: "searxng", url: "https://search.firma.local" }),
        Policy.parse({ allowedHosts: ["127.0.0.1", "search.firma.local"] }),
        isUrlAllowed,
      ),
    ).toBeUndefined();
  });

  it("organisationens politik kan låse søgetjenesten", () => {
    const settings = Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: "http://127.0.0.1:1/v1" }], activeEndpoint: "x" });
    expect(effectiveSearch(settings, Policy.parse({}))).toEqual({ provider: "duckduckgo" });
    const locked = Policy.parse({ search: { provider: "searxng", url: "https://search.firma.local" } });
    expect(effectiveSearch(settings, locked)).toEqual({ provider: "searxng", url: "https://search.firma.local" });
  });
});

describe("DuckDuckGo: ændret markup opdages", () => {
  it("resultater, som parseren ikke kan læse, skelnes fra 'ingen resultater'", () => {
    const changed =
      '<div class="results"><div class="result results_links"><h2 class="result__title"><a class="new-link-class" href="x">Titel</a></h2></div></div>';
    expect(parseDdgHtml(changed)).toEqual([]);
    expect(looksLikeUnparsedResults(changed)).toBe(true);
  });
  it("en ægte tom resultatside er ikke en fejl", () => {
    expect(looksLikeUnparsedResults('<div class="no-results">No results.</div>')).toBe(false);
    expect(looksLikeUnparsedResults("<html><body>Intet her</body></html>")).toBe(false);
  });
});

describe("DuckDuckGo: robot-tjek er ikke 'ingen resultater'", () => {
  // Uddrag af siden, DuckDuckGo svarer med (HTTP 202), når den tror, søgningen kommer fra en robot.
  const challenge =
    '<form id="challenge-form" action="//duckduckgo.com/anomaly.js?sv=html&cc=botnet"><div class="anomaly-modal__title">Unfortunately, bots use DuckDuckGo too.</div></form>';

  it("genkender siden med opgaven", () => {
    expect(parseDdgHtml(challenge)).toEqual([]);
    expect(looksLikeBotChallenge(challenge)).toBe(true);
    expect(looksLikeBotChallenge('<div class="no-results">No results.</div>')).toBe(false);
  });

  it("søgningen meldes som mislykket, og modellen får besked om ikke at gætte", async () => {
    const realFetch = globalThis.fetch;
    globalThis.fetch = (async () => new Response(challenge, { status: 202, headers: { "content-type": "text/html" } })) as typeof fetch;
    try {
      const r = await searchWeb({ provider: "duckduckgo" }, "novo nordisk aktiekurs", allowAll);
      expect(r.ok).toBe(false);
      expect(r.text).toBe(BOT_CHALLENGE_TEXT);
      expect(r.text).toContain("gæt ikke");
    } finally {
      globalThis.fetch = realFetch;
    }
  });
});
