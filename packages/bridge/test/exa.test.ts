import { afterEach, describe, expect, it } from "vitest";
import { Policy, Settings, checkSettings, effectiveSearch, isUrlAllowed } from "@miai/shared";
import { EXA_HOST, EXA_LIMIT_TEXT, searchWeb } from "../src/search.js";

/** Web-søgning gennem Exas hostede MCP-server (uden konto og nøgle). Serveren er simuleret med fetch. */

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

interface Seen {
  url: string;
  headers: Record<string, string>;
  body: { method: string; params: { name: string; arguments: { query: string; numResults: number } } };
}

/** Lader fetch svare som MCP-serveren og husker, hvad der blev sendt. */
function mockExa(reply: () => Response): Seen[] {
  const seen: Seen[] = [];
  globalThis.fetch = (async (url: string | URL | Request, init?: RequestInit) => {
    seen.push({ url: String(url), headers: init?.headers as Record<string, string>, body: JSON.parse(String(init?.body)) as Seen["body"] });
    return reply();
  }) as typeof fetch;
  return seen;
}

const sse = (payload: unknown) => new Response(`event: message\ndata: ${JSON.stringify(payload)}\n\n`, { headers: { "content-type": "text/event-stream" } });
const RESULT = "Title: Novo Nordisk – kurshistorik\nURL: https://eksempel.dk/novo\nHighlights:\n| 20. nov. 2026 | 264,50 |";
const allowAll = () => true;
const exa = { provider: "exa" as const };

describe("Exa som søgetjeneste", () => {
  it("kalder web_search_exa på MCP-serveren og giver modellen sidens uddrag", async () => {
    const seen = mockExa(() => sse({ jsonrpc: "2.0", id: 1, result: { content: [{ type: "text", text: RESULT }] } }));
    const r = await searchWeb(exa, "novo nordisk lukkekurs", allowAll);
    expect(r.ok).toBe(true);
    expect(r.host).toBe(EXA_HOST); // audit-loggen får den faktiske søgehost
    expect(r.text).toContain("Søgeresultater fra webbet");
    expect(r.text).toContain("https://eksempel.dk/novo");
    expect(r.text).toContain("264,50");

    expect(seen).toHaveLength(1);
    expect(seen[0].url).toBe("https://mcp.exa.ai/mcp");
    expect(seen[0].body).toMatchObject({ method: "tools/call", params: { name: "web_search_exa", arguments: { query: "novo nordisk lukkekurs" } } });
    // Uden konto: ingen nøgle og intet, der identificerer brugeren.
    expect(Object.keys(seen[0].headers).map((h) => h.toLowerCase())).toEqual(["content-type", "accept"]);
  });

  it("læser også et svar som almindelig JSON og tåler beskeder før resultatet i en strøm", async () => {
    mockExa(() => Response.json({ jsonrpc: "2.0", id: 1, result: { content: [{ type: "text", text: RESULT }] } }));
    expect((await searchWeb(exa, "x", allowAll)).text).toContain("264,50");

    const stream =
      `event: message\ndata: ${JSON.stringify({ jsonrpc: "2.0", method: "notifications/progress", params: {} })}\n\n` +
      `event: message\ndata: ${JSON.stringify({ jsonrpc: "2.0", id: 1, result: { content: [{ type: "text", text: RESULT }] } })}\n\n`;
    mockExa(() => new Response(stream, { headers: { "content-type": "text/event-stream" } }));
    expect((await searchWeb(exa, "x", allowAll)).text).toContain("264,50");
  });

  it("lange svar afkortes, så flere søgninger ikke fylder modellens kontekst", async () => {
    mockExa(() => sse({ jsonrpc: "2.0", id: 1, result: { content: [{ type: "text", text: "x".repeat(50_000) }] } }));
    expect((await searchWeb(exa, "x", allowAll)).text.length).toBe(8_000);
  });

  it("den gratis grænse (429) meldes som en fejl, og modellen får besked om ikke at gætte", async () => {
    mockExa(() => new Response("Too Many Requests", { status: 429 }));
    const r = await searchWeb(exa, "x", allowAll);
    expect(r).toEqual({ host: EXA_HOST, ok: false, text: EXA_LIMIT_TEXT });
    expect(r.text).toContain("gæt ikke");

    // Grænsen kan også komme som en fejl i selve svaret.
    mockExa(() => sse({ jsonrpc: "2.0", id: 1, result: { isError: true, content: [{ type: "text", text: "Rate limit exceeded for free tier" }] } }));
    expect((await searchWeb(exa, "x", allowAll)).text).toBe(EXA_LIMIT_TEXT);
  });

  it("andre fejl er ikke 'ingen resultater'", async () => {
    mockExa(() => sse({ jsonrpc: "2.0", id: 1, error: { code: -32602, message: "Unknown tool" } }));
    const rpc = await searchWeb(exa, "x", allowAll);
    expect(rpc.ok).toBe(false);
    expect(rpc.text).toContain("Unknown tool");

    mockExa(() => new Response("oops", { status: 503 }));
    expect(await searchWeb(exa, "x", allowAll)).toMatchObject({ ok: false, text: expect.stringContaining("503") });

    mockExa(() => new Response("<html>ikke et svar</html>"));
    expect(await searchWeb(exa, "x", allowAll)).toMatchObject({ ok: false, text: "Søgetjenesten gav et svar, der ikke kunne læses." });

    globalThis.fetch = (async () => {
      throw new Error("getaddrinfo ENOTFOUND mcp.exa.ai");
    }) as typeof fetch;
    expect(await searchWeb(exa, "x", allowAll)).toMatchObject({ host: EXA_HOST, ok: false, text: expect.stringContaining("svarede ikke") });
  });

  it("et tomt svar er 'ingen resultater'", async () => {
    mockExa(() => sse({ jsonrpc: "2.0", id: 1, result: { content: [] } }));
    expect(await searchWeb(exa, "x", allowAll)).toEqual({ host: EXA_HOST, ok: true, text: "Ingen resultater fra søgetjenesten." });
  });

  it("kan vælges i indstillingerne uden adresse og låses af organisationens politik", () => {
    const settings = Settings.parse({
      endpoints: [{ id: "x", name: "X", baseUrl: "http://127.0.0.1:1/v1" }],
      activeEndpoint: "x",
      search: { provider: "exa" },
    });
    expect(settings.search).toEqual({ provider: "exa" });
    expect(checkSettings(settings, Policy.parse({}), isUrlAllowed)).toBeUndefined();
    expect(effectiveSearch(Settings.parse({ ...settings, search: { provider: "duckduckgo" } }), Policy.parse({ search: { provider: "exa" } }))).toEqual({
      provider: "exa",
    });
  });
});
