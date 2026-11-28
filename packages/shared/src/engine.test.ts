import { createServer, type Server } from "node:http";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { executeTask, InputError, probeEndpoint, validateInput } from "./engine.js";
import { getTask } from "./tasks.js";
import { Endpoint, Policy, Settings, TaskRequest, type TaskEvent } from "./types.js";

// Simulerer et inference-værktøj, som add-in'et kalder direkte (som Ollama med CORS).
let server: Server;
let base: string;
const seen: { method: string; headers: Record<string, unknown> }[] = [];
let lastBody: Record<string, unknown> = {};

beforeAll(async () => {
  server = createServer((req, res) => {
    seen.push({ method: req.method!, headers: req.headers });
    res.setHeader("access-control-allow-origin", "*");
    if (req.url === "/v1/models") return res.end(JSON.stringify({ data: [{ id: "gemma3:4b" }] }));
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      if (body) lastBody = JSON.parse(body);
      // Server der ignorerer stream=true og svarer med ét JSON-objekt.
      res.setHeader("content-type", "application/json");
      // TOOLTEST: simulerer en model, der kalder web_search. Uden tools-parameter (sidste runde) svarer den normalt.
      if (body.includes("TOOLTEST")) {
        let wants: { messages?: { role?: string }[]; tools?: unknown };
        try {
          wants = JSON.parse(body);
        } catch {
          wants = {};
        }
        if (!wants.tools) {
          res.end(JSON.stringify({ choices: [{ message: { content: "Endeligt svar uden værktøjer." } }] }));
          return;
        }
        if (!wants.messages?.some((m) => m.role === "tool")) {
          res.end(
            JSON.stringify({
              choices: [
                {
                  finish_reason: "tool_calls",
                  message: {
                    content: "",
                    tool_calls: [{ id: "call_1", type: "function", function: { name: "web_search", arguments: '{"query":"test query"}' } }],
                  },
                },
              ],
            }),
          );
          return;
        }
        res.end(JSON.stringify({ choices: [{ message: { content: "Endeligt svar baseret på søgning." } }] }));
        return;
      }
      // LOOPTEST: simulerer en model, der er faldet i gentagelsessløjfe.
      const content = body.includes("LOOPTEST") ? `Jeg kigger på indholdet. ${Array(8).fill("side 6 tabel").join(" ")}` : "Kort svar.";
      res.end(JSON.stringify({ choices: [{ message: { content } }] }));
    });
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  base = `http://127.0.0.1:${(server.address() as { port: number }).port}/v1`;
});
afterAll(() => server.close());

const settings = (ep: object = {}, extra: object = {}) =>
  Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base, ...ep }], activeEndpoint: "x", ...extra });

async function run(id: string, body: object, policy = Policy.parse({}), hint?: string, ep: object = {}) {
  const events: TaskEvent[] = [];
  const s = settings(ep);
  const req = TaskRequest.parse(body);
  await executeTask(getTask(id)!, req, {
    settings: s,
    endpoint: s.endpoints[0],
    policy,
    emit: (e) => events.push(e),
    signal: new AbortController().signal,
    client: { unreachableHint: hint },
  });
  return events;
}

describe("executeTask (direkte tilstand)", () => {
  it("vælger første model, filtrerer <think> og tåler ikke-streamende server", async () => {
    const ev = await run("simplify", { context: { host: "word", text: "Lang tekst" } });
    expect(ev[0]).toMatchObject({ type: "meta", model: "gemma3:4b" });
    expect(
      ev
        .filter((e) => e.type === "delta")
        .map((e: any) => e.text)
        .join(""),
    ).toBe("Kort svar.");
  });

  it("sender ingen ekstra headers på GET (undgår CORS-preflight)", () => {
    const get = seen.find((s) => s.method === "GET")!;
    expect(get.headers["content-type"]).toBeUndefined();
    expect(get.headers.authorization).toBeUndefined();
  });

  it("håndhæver politikken", async () => {
    await expect(run("simplify", { context: { host: "word", text: "x" } }, Policy.parse({ allowedHosts: ["llm.firma.local"] }))).rejects.toThrow(/politikken/);
  });

  it("giver hint ved forbindelsesfejl", async () => {
    const s = Settings.parse({ endpoints: [{ id: "d", name: "Død", baseUrl: "http://127.0.0.1:9/v1" }], activeEndpoint: "d", defaultModel: "m" });
    const p = executeTask(getTask("simplify")!, TaskRequest.parse({ context: { host: "word", text: "x" } }), {
      settings: s,
      endpoint: s.endpoints[0],
      policy: Policy.parse({}),
      emit: () => {},
      signal: new AbortController().signal,
      client: { unreachableHint: "Tjek CORS." },
    });
    await expect(p).rejects.toThrow(/Tjek CORS/);
  });

  it("validerer input", () => {
    expect(() => validateInput(getTask("explain-formula")!, TaskRequest.parse({ context: { host: "excel" } }))).toThrow(InputError);
  });

  it("afbryder ved gentagelsessløjfe og markerer svaret", async () => {
    const ev = await run("chat", { context: { host: "word" }, params: { prompt: "LOOPTEST" } });
    const text = ev
      .filter((e) => e.type === "delta")
      .map((e: { text?: string }) => e.text ?? "")
      .join("");
    expect(text).not.toContain("gentage sig selv"); // ingen note i selve teksten
    expect(ev.find((e) => e.type === "done")).toMatchObject({ stopReason: "repetition" });
    expect(ev.filter((e) => e.type === "error").length).toBe(0);
  });

  it("sender presence_penalty til fritekst, når den er sat", async () => {
    await run("chat", { context: { host: "word" }, params: { prompt: "hej" } }, undefined, undefined, { presencePenalty: 1.5 });
    expect(lastBody.presence_penalty).toBe(1.5);
  });

  it("udelader presence_penalty for JSON-opgaver og Excel-celler (TSV)", async () => {
    const ep = { presencePenalty: 1.5, structuredOutput: "none" };
    await run("generate-formula", { context: { host: "excel", address: "A1" }, params: { prompt: "sum" } }, undefined, undefined, ep).catch(() => {});
    expect(lastBody).not.toHaveProperty("presence_penalty");
    await run("translate", { context: { host: "excel", text: "Ja", table: "Ja\nJa" } }, undefined, undefined, ep);
    expect(lastBody).not.toHaveProperty("presence_penalty");
  });

  it("sender ikke presence_penalty, når den ikke er sat", async () => {
    await run("chat", { context: { host: "word" }, params: { prompt: "hej" } });
    expect(lastBody).not.toHaveProperty("presence_penalty");
  });

  it("inkluderer arkoversigt og områder i Excel-chat (mod partielle svar som '30 for hele kolonnen')", async () => {
    await run("chat", {
      context: {
        host: "excel",
        address: "Ark1!E9:E30",
        table: "5\n6",
        surroundings:
          "Active sheet 'Ark1', used range A1:E30 (30 rows x 5 columns)\nHeaders (row 1): E=Beløb\nColumn stats (30 rows): E: 29 non-empty, numeric sum=1234 min=10 max=98",
      },
      params: { prompt: "Hvad summer kolonne E op til?" },
    });
    const sys = lastBody.messages?.[0].content as string;
    expect(sys).toContain("Column stats (30 rows)"); // oversigten når til modellen
    expect(sys).toContain("Never present figures from the selection as if they covered the whole sheet");
    expect(sys).toContain("Ark1!E9:E30"); // markeringen er navngivet
  });

  it("kører tool-løkke (web_search) og udsender search-event", async () => {
    const calls: string[] = [];
    const events: TaskEvent[] = [];
    const s = settings({}, { webSearch: true });
    await executeTask(getTask("chat")!, TaskRequest.parse({ context: { host: "word" }, params: { prompt: "TOOLTEST" } }), {
      settings: s,
      endpoint: s.endpoints[0],
      policy: Policy.parse({}),
      emit: (e) => events.push(e),
      signal: new AbortController().signal,
      now: () => new Date(2026, 10, 25, 23, 30),
      tools: {
        specs: [{ type: "function", function: { name: "web_search", description: "Søg på webbet", parameters: {} } }],
        run: async (_name, args) => {
          calls.push(args);
          return "1. Titel\n   https://eksempel.dk\n   Snippet fra siden.";
        },
      },
    });
    expect(calls).toHaveLength(1); // præcis ét værktøjskald, derefter endeligt svar
    // Med søgning kender modellen dags dato (ellers afviser den nylige datoer som fremtidige).
    expect(lastBody.messages?.[0].content).toContain("Today's date is 2026-11-25.");
    expect(JSON.parse(calls[0])).toEqual({ query: "test query" });
    const text = events
      .filter((e) => e.type === "delta")
      .map((e: { text?: string }) => e.text ?? "")
      .join("");
    expect(text).toContain("Endeligt svar baseret på søgning.");
    expect(events.some((e) => e.type === "search" && (e as { query: string }).query === "test query")).toBe(true);
    // Værktøjskaldet sendes tilbage med type: "function" (kræves af OpenAI-specifikationen, fx vLLM).
    const assistant = (lastBody.messages as { role: string; tool_calls?: { type?: string }[] }[]).find((m) => m.tool_calls);
    expect(assistant?.tool_calls?.[0].type).toBe("function");
  });

  it("web-søgning er slået fra som standard og kan forbydes af politikken", async () => {
    const calls: string[] = [];
    const tools = {
      specs: [{ type: "function" as const, function: { name: "web_search", description: "", parameters: {} } }],
      run: async () => (calls.push("x"), "resultat"),
    };
    const req = TaskRequest.parse({ context: { host: "word" }, params: { prompt: "TOOLTEST" } });
    const base = { emit: () => {}, signal: new AbortController().signal, tools };
    const def = settings();
    expect(def.webSearch).toBe(false);
    await executeTask(getTask("chat")!, req, { ...base, settings: def, endpoint: def.endpoints[0], policy: Policy.parse({}) });
    const on = settings({}, { webSearch: true });
    await executeTask(getTask("chat")!, req, { ...base, settings: on, endpoint: on.endpoints[0], policy: Policy.parse({ webSearch: false }) });
    expect(calls).toHaveLength(0);
  });

  it("kører ikke værktøjer, når webSearch er slået fra", async () => {
    const calls: string[] = [];
    const s = Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base }], activeEndpoint: "x", webSearch: false });
    await executeTask(getTask("chat")!, TaskRequest.parse({ context: { host: "word" }, params: { prompt: "TOOLTEST" } }), {
      settings: s,
      endpoint: s.endpoints[0],
      policy: Policy.parse({}),
      emit: () => {},
      signal: new AbortController().signal,
      tools: { specs: [], run: async () => (calls.push("x"), "resultat") },
    });
    expect(calls).toHaveLength(0); // eksekvereren må aldrig kaldes
  });
});

describe("probeEndpoint: lokal netværksadgang (#74)", () => {
  // Intet lytter på port 9, så kaldet fejler som en lukket port – ligesom et kald, browseren har blokeret.
  const dead = Endpoint.parse({ id: "d", name: "Død", baseUrl: "http://127.0.0.1:9/v1" });
  const policy = Policy.parse({});

  it("afvist tilladelse giver blocked – og der prøves ikke med no-cors", async () => {
    for (const lna of ["denied", "frame"] as const) {
      let noCors = 0;
      const r = await probeEndpoint(dead, policy, { localNetworkAccess: async () => lna, reachableWithoutCors: async () => (noCors++, true) });
      expect(r).toMatchObject({ ok: false, reason: "blocked", lna });
      expect(!r.ok && r.message).toContain("http://127.0.0.1:9/v1");
      expect(noCors).toBe(0);
    }
  });

  it("uden svar fra tilladelsen er resultatet som før: cors, når noget lytter, ellers absent", async () => {
    const unknown = async () => undefined;
    expect(await probeEndpoint(dead, policy, { localNetworkAccess: unknown, reachableWithoutCors: async () => true })).toMatchObject({ reason: "cors" });
    const absent = await probeEndpoint(dead, policy, { localNetworkAccess: unknown, reachableWithoutCors: async () => false });
    expect(absent).toMatchObject({ ok: false, reason: "absent" });
    expect(absent).not.toHaveProperty("lna");
    expect(await probeEndpoint(dead, policy)).toMatchObject({ ok: false, reason: "absent" });
  });

  it("ubesvaret tilladelse følger med et absent, og en fejl i opslaget ignoreres", async () => {
    expect(await probeEndpoint(dead, policy, { localNetworkAccess: async () => "prompt" })).toMatchObject({ reason: "absent", lna: "prompt" });
    const failing = () => Promise.reject(new Error("ingen Permissions API"));
    expect(await probeEndpoint(dead, policy, { localNetworkAccess: failing })).toMatchObject({ ok: false, reason: "absent" });
  });

  it("en endpoint, der svarer, spørger aldrig om tilladelsen", async () => {
    let asked = 0;
    const r = await probeEndpoint(Endpoint.parse({ id: "x", name: "X", baseUrl: base }), policy, { localNetworkAccess: async () => (asked++, "denied") });
    expect(r).toEqual({ ok: true, models: ["gemma3:4b"] });
    expect(asked).toBe(0);
  });
});
