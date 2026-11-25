import { createServer, type Server } from "node:http";
import type { AddressInfo } from "node:net";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { clearModelCache, executeTask, modelDetails } from "./engine.js";
import { fetchNativeModelInfo, parseLmStudioModels, parseOllamaPs, parseOllamaShow, serverRoot } from "./modelInfo.js";
import { getTask } from "./tasks.js";
import { Endpoint, Policy, Settings, TaskRequest, type TaskEvent } from "./types.js";

/** Modellens kontekst og kapabiliteter fra Ollamas og LM Studios eget API (#73). */

const ollamaShow = (parameters = "") => ({
  parameters,
  model_info: { "general.architecture": "qwen3", "qwen3.context_length": 40_960 },
  capabilities: ["completion", "tools", "thinking"],
});

const lmStudioModels = {
  models: [
    {
      type: "llm",
      key: "qwen/qwen3-8b",
      max_context_length: 32_768,
      loaded_instances: [{ id: "qwen/qwen3-8b", config: { context_length: 4_096 } }],
      capabilities: { vision: false, trained_for_tool_use: true },
    },
    { type: "llm", key: "google/gemma-3-12b", max_context_length: 131_072, loaded_instances: [], capabilities: { vision: true, trained_for_tool_use: false } },
  ],
};

describe("aflæsning af værktøjernes svar", () => {
  it("Ollama /api/show: num_ctx fra Modelfile, trænet kontekst og kapabiliteter", () => {
    expect(parseOllamaShow("qwen3:8b", ollamaShow('num_ctx                        8192\nstop "<|im_end|>"'))).toEqual({
      model: "qwen3:8b",
      server: "ollama",
      contextTokens: 8_192,
      contextSource: "ollama-modelfile",
      maxContextTokens: 40_960,
      vision: false,
      tools: true,
      thinking: true,
    });
    // Uden num_ctx kendes konteksten først, når modellen er indlæst.
    expect(parseOllamaShow("qwen3:8b", ollamaShow())?.contextTokens).toBeUndefined();
    expect(parseOllamaShow("x", { error: "model 'x' not found" })).toBeUndefined();
    expect(parseOllamaShow("x", "Kort svar.")).toBeUndefined();
  });

  it("Ollama /api/ps: den indlæste kontekst, også uden :latest", () => {
    const ps = { models: [{ name: "qwen3:latest", model: "qwen3:latest", context_length: 4_096 }] };
    expect(parseOllamaPs("qwen3", ps)).toBe(4_096);
    expect(parseOllamaPs("qwen3:latest", ps)).toBe(4_096);
    expect(parseOllamaPs("qwen3:8b", ps)).toBeUndefined();
    expect(parseOllamaPs("qwen3", { models: [] })).toBeUndefined();
  });

  it("LM Studio /api/v1/models: indlæst kontekst, maks., vision og værktøjer", () => {
    expect(parseLmStudioModels("qwen/qwen3-8b", lmStudioModels)).toEqual({
      model: "qwen/qwen3-8b",
      server: "lmstudio",
      contextTokens: 4_096,
      contextSource: "lmstudio-loaded",
      maxContextTokens: 32_768,
      vision: false,
      tools: true,
    });
    // Ikke indlæst: kun maks. og kapabiliteter.
    const unloaded = parseLmStudioModels("google/gemma-3-12b", lmStudioModels);
    expect(unloaded).toMatchObject({ maxContextTokens: 131_072, vision: true });
    expect(unloaded?.contextTokens).toBeUndefined();
    expect(parseLmStudioModels("m", { data: [] })).toBeUndefined();
  });

  it("rod-URL uden /v1", () => {
    expect(serverRoot("http://127.0.0.1:11434/v1")).toBe("http://127.0.0.1:11434");
    expect(serverRoot("http://127.0.0.1:1234/v1/")).toBe("http://127.0.0.1:1234");
  });
});

// Én server, der efter behov opfører sig som Ollama, LM Studio eller et andet værktøj (fx vLLM).
let server: Server;
let base = "";
let kind: "ollama" | "lmstudio" | "other" = "ollama";
let parameters = "";
let loaded: number | undefined = 4_096;
/** Ollama indlæser modellen ved første chat-kald (så /api/ps kender konteksten bagefter). */
let loadOnChat = false;
const hits: string[] = [];

beforeAll(async () => {
  server = createServer((req, res) => {
    hits.push(`${req.method} ${req.url}`);
    let raw = "";
    req.on("data", (c) => (raw += c));
    req.on("end", () => {
      res.setHeader("content-type", "application/json");
      if (req.url === "/v1/models") return res.end(JSON.stringify({ data: [{ id: "qwen3:8b" }] }));
      if (kind === "ollama" && req.url === "/api/show") return res.end(JSON.stringify(ollamaShow(parameters)));
      if (kind === "ollama" && req.url === "/api/ps")
        return res.end(JSON.stringify({ models: loaded ? [{ name: "qwen3:8b", model: "qwen3:8b", context_length: loaded }] : [] }));
      if (kind === "lmstudio" && req.url === "/api/v1/models") return res.end(JSON.stringify(lmStudioModels));
      if (req.url !== "/v1/chat/completions") return res.writeHead(404).end();
      if (loadOnChat) loaded = 4_096;
      res.end(JSON.stringify({ choices: [{ message: { content: "Kort svar." }, finish_reason: "stop" }] }));
    });
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`;
});
afterAll(() => new Promise<void>((r) => server.close(() => r())));
beforeEach(() => {
  clearModelCache();
  kind = "ollama";
  parameters = "";
  loaded = 4_096;
  loadOnChat = false;
  hits.length = 0;
});

const endpoint = (extra: Record<string, unknown> = {}) => Endpoint.parse({ id: "e", name: "E", baseUrl: base, ...extra });

describe("modelDetails", () => {
  it("Ollama: den indlæste kontekst går forud for Modelfile", async () => {
    parameters = "num_ctx 8192";
    expect(await modelDetails(endpoint(), "qwen3:8b")).toMatchObject({ server: "ollama", contextTokens: 4_096, contextSource: "ollama-loaded", tools: true });
    loaded = undefined;
    expect(await modelDetails(endpoint(), "qwen3:8b", { fresh: true })).toMatchObject({ contextTokens: 8_192, contextSource: "ollama-modelfile" });
  });

  it("Ollama uden num_ctx og ikke indlæst: ukendt kontekst, men værktøjet er genkendt", async () => {
    loaded = undefined;
    const d = await modelDetails(endpoint(), "qwen3:8b");
    expect(d.server).toBe("ollama");
    expect(d.contextTokens).toBeUndefined();
    expect(d.maxContextTokens).toBe(40_960);
  });

  it("LM Studio på sin standardport spørges først; ellers efter Ollama", async () => {
    kind = "lmstudio";
    expect(await modelDetails(endpoint(), "qwen/qwen3-8b")).toMatchObject({ server: "lmstudio", contextTokens: 4_096, maxContextTokens: 32_768 });
    expect(hits).toContain("POST /api/show"); // ikke port 1234: Ollama prøves først
  });

  it("endpointens kontekstlængde har forrang", async () => {
    expect(await modelDetails(endpoint({ contextTokens: 65_536 }), "qwen3:8b")).toMatchObject({
      contextTokens: 65_536,
      contextSource: "endpoint",
      server: "ollama",
    });
  });

  it("et andet værktøj spørges ikke igen i hver kørsel", async () => {
    kind = "other";
    expect(await fetchNativeModelInfo(endpoint(), "qwen3:8b")).toBeUndefined();
    const n = hits.length;
    expect(await fetchNativeModelInfo(endpoint(), "qwen3:8b")).toBeUndefined();
    expect(hits.length).toBe(n);
    expect(await fetchNativeModelInfo(endpoint(), "qwen3:8b", { fresh: true })).toBeUndefined();
    expect(hits.length).toBeGreaterThan(n);
  });
});

describe("motoren med lille kontekst", () => {
  const run = async (taskId: string, text: string) => {
    const settings = Settings.parse({ endpoints: [endpoint()], activeEndpoint: "e" });
    const events: TaskEvent[] = [];
    let error: unknown;
    try {
      await executeTask(getTask(taskId)!, TaskRequest.parse({ context: { host: "word", text }, params: { prompt: "Hvad står der?" } }), {
        settings,
        endpoint: settings.endpoints[0],
        policy: Policy.parse({}),
        emit: (e) => events.push(e),
        signal: new AbortController().signal,
      });
    } catch (e) {
      error = e;
    }
    return { done: events.find((e) => e.type === "done") as Extract<TaskEvent, { type: "done" }> | undefined, error };
  };

  it("chat på Ollama med 4.096 tokens: advarsel med værktøjet", async () => {
    const { done } = await run("chat", "Kort tekst.");
    expect(done?.warnings).toContain("small-context");
    expect(done?.context).toEqual({ tokens: 4_096, server: "ollama" });
  });

  it("advarslen kommer også, når modellen først blev indlæst af kørslen", async () => {
    loaded = undefined;
    loadOnChat = true;
    const { done } = await run("chat", "Kort tekst.");
    expect(done?.context).toEqual({ tokens: 4_096, server: "ollama" });
  });

  it("andre opgaver advares ikke, og stor kontekst giver ingen advarsel", async () => {
    expect((await run("simplify", "Kort tekst.")).done?.warnings).toBeUndefined();
    clearModelCache();
    loaded = 32_768;
    expect((await run("chat", "Kort tekst.")).done?.warnings).toBeUndefined();
  });

  it("en prompt over den kendte kontekst afvises før afsendelse, med rettelsen til Ollama", async () => {
    const { error } = await run("simplify", "Et afsnit med noget tekst. ".repeat(1_500));
    expect((error as { code?: string }).code).toBe("CONTEXT_OVERFLOW");
    expect((error as Error).message).toContain("OLLAMA_CONTEXT_LENGTH");
    expect(hits).not.toContain("POST /v1/chat/completions");
  });
});
