import { createServer, type IncomingMessage, type Server } from "node:http";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { FastifyInstance } from "fastify";
import { extractJson, isUrlAllowed, Settings, ThinkFilter, type TaskEvent } from "@miai/shared";
import { ConfigStore } from "../src/config.js";
import { buildServer } from "../src/server.js";

// ---------------------------------------------------------------------------
// Mock af en OpenAI-kompatibel inference-server
// ---------------------------------------------------------------------------
let mock: Server;
let mockUrl: string;
let nextReply: string[] = [];
let lastBody: any;

async function readBody(req: IncomingMessage): Promise<string> {
  let s = "";
  for await (const c of req) s += c;
  return s;
}

beforeAll(async () => {
  mock = createServer(async (req, res) => {
    if (req.url === "/v1/models") {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ object: "list", data: [{ id: "mock-model" }, { id: "other" }] }));
    }
    if (req.url === "/v1/chat/completions") {
      lastBody = JSON.parse(await readBody(req));
      const parts = nextReply.shift()?.split("|") ?? ["ok"];
      res.writeHead(200, { "content-type": "text/event-stream" });
      for (const p of parts) res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: p } }] })}\n\n`);
      res.write(`data: ${JSON.stringify({ choices: [{ delta: {} }], usage: { prompt_tokens: 10, completion_tokens: 5 } })}\n\n`);
      return res.end("data: [DONE]\n\n");
    }
    res.statusCode = 404;
    res.end();
  });
  await new Promise<void>((r) => mock.listen(0, "127.0.0.1", r));
  const addr = mock.address() as { port: number };
  mockUrl = `http://127.0.0.1:${addr.port}/v1`;
});

afterAll(() => mock.close());

const PORT = 44321;
const H = { host: `localhost:${PORT}`, "x-miai": "1" };
let app: FastifyInstance;
let dir: string;

beforeEach(async () => {
  dir = mkdtempSync(join(tmpdir(), "miai-"));
  const store = new ConfigStore(dir, join(dir, "no-policy.json"));
  store.update(Settings.parse({ endpoints: [{ id: "mock", name: "Mock", baseUrl: mockUrl }], activeEndpoint: "mock" }));
  app = await buildServer({ store, port: PORT });
});

function events(payload: string): TaskEvent[] {
  return payload
    .split("\n\n")
    .filter((b) => b.startsWith("data: "))
    .map((b) => JSON.parse(b.slice(6)));
}

describe("sikkerhed", () => {
  it("afviser fremmed Host (DNS rebinding)", async () => {
    const r = await app.inject({ url: "/api/health", headers: { host: "evil.example:44321", "x-miai": "1" } });
    expect(r.statusCode).toBe(421);
  });
  it("afviser fremmed origin og manglende X-MIAI", async () => {
    expect((await app.inject({ url: "/api/health", headers: { ...H, origin: "https://evil.example" } })).statusCode).toBe(403);
    expect((await app.inject({ url: "/api/health", headers: { host: H.host } })).statusCode).toBe(403);
  });
  it("tillader egen origin", async () => {
    const r = await app.inject({ url: "/api/health", headers: { ...H, origin: `https://localhost:${PORT}` } });
    expect(r.json()).toMatchObject({ ok: true, endpoint: { reachable: true } });
  });
});

describe("politik", () => {
  it("matcher hosts og wildcards", () => {
    const p = { allowedHosts: ["127.0.0.1", "localhost", "::1", "*.firma.local"], locked: false };
    expect(isUrlAllowed("http://127.0.0.1:11434/v1", p)).toBe(true);
    expect(isUrlAllowed("http://[::1]:8080/v1", p)).toBe(true);
    expect(isUrlAllowed("https://llm.firma.local/v1", p)).toBe(true);
    expect(isUrlAllowed("https://firma.local.evil.com/v1", p)).toBe(false);
    expect(isUrlAllowed("https://api.openai.com/v1", p)).toBe(false);
  });
  it("afviser endpoint uden for politikken", async () => {
    const r = await app.inject({
      method: "PUT",
      url: "/api/settings",
      headers: H,
      payload: { settings: { endpoints: [{ id: "x", name: "Cloud", baseUrl: "https://api.example.com/v1" }], activeEndpoint: "x" } },
    });
    expect(r.statusCode).toBe(400);
    expect(r.json().error).toContain("ikke tilladt");
  });
  it("maskinpolitik er låst og har forrang", () => {
    const pf = join(dir, "policy.json");
    writeFileSync(pf, JSON.stringify({ allowedHosts: ["llm.firma.local"], locked: true }));
    const store = new ConfigStore(dir, pf);
    expect(store.policySource).toBe("machine");
    expect(() => store.update(store.getSettings(), { allowedHosts: ["*"], locked: false })).toThrow(/låst/);
  });
  it("ulæselig maskinpolitik lukker ned (fail closed)", () => {
    const pf = join(dir, "bad.json");
    writeFileSync(pf, "{ ikke json");
    const store = new ConfigStore(dir, pf);
    expect(store.getPolicy()).toMatchObject({ locked: true, allowedHosts: ["127.0.0.1", "localhost", "::1"] });
  });
  it("maskerer API-nøgler og bevarer dem ved gem", async () => {
    const settings = { endpoints: [{ id: "mock", name: "Mock", baseUrl: mockUrl, apiKey: "hemmelig" }], activeEndpoint: "mock" };
    await app.inject({ method: "PUT", url: "/api/settings", headers: H, payload: { settings } });
    const view = (await app.inject({ url: "/api/settings", headers: H })).json();
    expect(view.settings.endpoints[0].apiKey).toBe("********");
    await app.inject({ method: "PUT", url: "/api/settings", headers: H, payload: { settings: view.settings } });
    nextReply = ["hej"];
    await app.inject({ method: "POST", url: "/api/tasks/simplify", headers: H, payload: { context: { host: "word", text: "x" } } });
    expect(lastBody).toBeDefined();
  });
});

describe("opgaver", () => {
  it("streamer tekst og filtrerer <think>", async () => {
    nextReply = ["<thi|nk>overvejer…</th|ink>\n\nEnkel| tekst."];
    const r = await app.inject({ method: "POST", url: "/api/tasks/simplify", headers: H, payload: { context: { host: "word", text: "Kompliceret tekst." } } });
    const ev = events(r.payload);
    expect(ev[0]).toMatchObject({ type: "meta", model: "mock-model" });
    const text = ev.filter((e) => e.type === "delta").map((e: any) => e.text).join("");
    expect(text).toBe("Enkel tekst.");
    expect(ev.at(-1)).toMatchObject({ type: "done", usage: { promptTokens: 10 } });
    expect(lastBody.stream).toBe(true);
    expect(lastBody.messages[1].content).toContain("<document>");
  });

  it("returnerer valideret JSON og sender json_schema", async () => {
    nextReply = ['```json\n{"edits":[{"original":"hun|d","replacement":"hund","explanation":"stavefejl"}]}\n```'];
    const r = await app.inject({ method: "POST", url: "/api/tasks/grammar", headers: H, payload: { context: { host: "word", text: "En hun d." } } });
    const result = events(r.payload).find((e) => e.type === "result") as any;
    expect(result.data.edits[0].replacement).toBe("hund");
    expect(lastBody.response_format.type).toBe("json_schema");
    expect(lastBody.temperature).toBe(0);
  });

  it("prøver igen ved ugyldig JSON", async () => {
    nextReply = ["ikke json", '{"formula":"=SUM(B:B)","explanation":"Summerer B"}'];
    const r = await app.inject({
      method: "POST",
      url: "/api/tasks/generate-formula",
      headers: H,
      payload: { context: { host: "excel", address: "C1" }, params: { prompt: "summen af B" } },
    });
    const result = events(r.payload).find((e) => e.type === "result") as any;
    expect(result.data.formula).toBe("=SUM(B:B)");
    expect(lastBody.messages.at(-1).content).toContain("not valid");
  });

  it("gemmer presencePenalty og sender den til fritekst", async () => {
    const settings = { endpoints: [{ id: "mock", name: "Mock", baseUrl: mockUrl, presencePenalty: 1 }], activeEndpoint: "mock" };
    expect((await app.inject({ method: "PUT", url: "/api/settings", headers: H, payload: { settings } })).statusCode).toBe(200);
    expect((await app.inject({ url: "/api/settings", headers: H })).json().settings.endpoints[0].presencePenalty).toBe(1);
    nextReply = ["ok"];
    await app.inject({ method: "POST", url: "/api/tasks/simplify", headers: H, payload: { context: { host: "word", text: "x" } } });
    expect(lastBody.presence_penalty).toBe(1);
  });

  it("kræver input", async () => {
    const r = await app.inject({ method: "POST", url: "/api/tasks/explain-formula", headers: H, payload: { context: { host: "excel" } } });
    expect(r.statusCode).toBe(400);
  });

  it("melder fejl når endpoint ikke svarer", async () => {
    const store = new ConfigStore(dir, join(dir, "none.json"));
    store.update(Settings.parse({ endpoints: [{ id: "dead", name: "Død", baseUrl: "http://127.0.0.1:9/v1" }], activeEndpoint: "dead", defaultModel: "m" }));
    const dead = await buildServer({ store, port: PORT });
    const r = await dead.inject({ method: "POST", url: "/api/tasks/simplify", headers: H, payload: { context: { host: "word", text: "x" } } });
    expect(events(r.payload).at(-1)).toMatchObject({ type: "error", code: "UPSTREAM_UNREACHABLE" });
  });
});

describe("hjælpere", () => {
  it("ThinkFilter håndterer split tags", () => {
    const f = new ThinkFilter();
    const out = ["a<", "think>x</", "think>b<t", "x"].map((c) => f.push(c)).join("") + f.flush();
    expect(out).toBe("ab<tx");
  });
  it("extractJson tåler tekst omkring", () => {
    expect(extractJson('Her er svaret: {"a":"}"} tak')).toEqual({ a: "}" });
  });
});
