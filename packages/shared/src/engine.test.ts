import { createServer, type Server } from "node:http";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { executeTask, InputError, validateInput } from "./engine.js";
import { getTask } from "./tasks.js";
import { Policy, Settings, TaskRequest, type TaskEvent } from "./types.js";

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
                { finish_reason: "tool_calls", message: { content: "", tool_calls: [{ id: "call_1", type: "function", function: { name: "web_search", arguments: '{"query":"test query"}' } }] } },
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

const settings = (ep: object = {}) => Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base, ...ep }], activeEndpoint: "x" });

async function run(id: string, body: object, policy = Policy.parse({}), hint?: string, ep: object = {}) {
  const events: TaskEvent[] = [];
  const s = settings(ep);
  const req = TaskRequest.parse(body);
  await executeTask(getTask(id)!, req, { settings: s, endpoint: s.endpoints[0], policy, emit: (e) => events.push(e), signal: new AbortController().signal, client: { unreachableHint: hint } });
  return events;
}

describe("executeTask (direkte tilstand)", () => {
  it("vælger første model, filtrerer <think> og tåler ikke-streamende server", async () => {
    const ev = await run("simplify", { context: { host: "word", text: "Lang tekst" } });
    expect(ev[0]).toMatchObject({ type: "meta", model: "gemma3:4b" });
    expect(ev.filter((e) => e.type === "delta").map((e: any) => e.text).join("")).toBe("Kort svar.");
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
    const text = ev.filter((e) => e.type === "delta").map((e: { text?: string }) => e.text ?? "").join("");
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

  it("kører tool-løkke (web_search) og udsender search-event", async () => {
    const calls: string[] = [];
    const events: TaskEvent[] = [];
    const s = settings();
    await executeTask(
      getTask("chat")!,
      TaskRequest.parse({ context: { host: "word" }, params: { prompt: "TOOLTEST" } }),
      {
        settings: s,
        endpoint: s.endpoints[0],
        policy: Policy.parse({}),
        emit: (e) => events.push(e),
        signal: new AbortController().signal,
        tools: {
          specs: [{ type: "function", function: { name: "web_search", description: "Søg på webbet", parameters: {} } }],
          run: async (_name, args) => {
            calls.push(args);
            return "1. Titel\n   https://eksempel.dk\n   Snippet fra siden.";
          },
        },
      },
    );
    expect(calls).toHaveLength(1); // præcis ét værktøjskald, derefter endeligt svar
    expect(JSON.parse(calls[0])).toEqual({ query: "test query" });
    const text = events.filter((e) => e.type === "delta").map((e: { text?: string }) => e.text ?? "").join("");
    expect(text).toContain("Endeligt svar baseret på søgning.");
    expect(events.some((e) => e.type === "search" && (e as { query: string }).query === "test query")).toBe(true);
  });

  it("kører ikke værktøjer, når webSearch er slået fra", async () => {
    const calls: string[] = [];
    const s = Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base }], activeEndpoint: "x", webSearch: false });
    await executeTask(
      getTask("chat")!,
      TaskRequest.parse({ context: { host: "word" }, params: { prompt: "TOOLTEST" } }),
      {
        settings: s,
        endpoint: s.endpoints[0],
        policy: Policy.parse({}),
        emit: () => {},
        signal: new AbortController().signal,
        tools: { specs: [], run: async () => (calls.push("x"), "resultat") },
      },
    );
    expect(calls).toHaveLength(0); // eksekvereren må aldrig kaldes
  });
});
