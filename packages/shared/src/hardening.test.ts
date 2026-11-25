import { createServer, type Server } from "node:http";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { chunkChars, estimateTokens, looksTruncated, outputBudget, splitIntoChunks } from "./context.js";
import { clearModelCache, errorEvent, executeTask } from "./engine.js";
import { UpstreamError } from "./openai.js";
import { mergeCellAudit, type AuditEntry } from "./audit.js";
import { modelGroupOf } from "./modelGroups.js";
import { getTask, tasksForHost } from "./tasks.js";
import { Policy, Settings, TaskRequest, type TaskEvent } from "./types.js";

/** Hærdning (#3): kontekstbudget, afkortning, opdeling af lange tekster, fejlkoder. */

interface Reply {
  text?: string;
  promptTokens?: number;
  finishReason?: string;
  toolCall?: boolean;
}

let server: Server;
let base: string;
let bodies: Record<string, unknown>[] = [];
let models: Record<string, unknown>[] = [{ id: "m" }];
let reply: (body: Record<string, unknown>) => Reply = () => ({ text: "ok" });
let rejectStreamOptions = false;

beforeAll(async () => {
  server = createServer((req, res) => {
    if (req.url === "/v1/models") return res.end(JSON.stringify({ data: models }));
    // Som et værktøj uden Ollamas eller LM Studios eget API (#73).
    if (!req.url?.endsWith("/chat/completions")) return res.writeHead(404).end();
    let raw = "";
    req.on("data", (c) => (raw += c));
    req.on("end", () => {
      const body = JSON.parse(raw);
      if (rejectStreamOptions && body.stream_options) {
        res.statusCode = 400;
        return res.end(JSON.stringify({ error: { message: "Unrecognized request argument supplied: stream_options" } }));
      }
      bodies.push(body);
      const r = reply(body);
      res.writeHead(200, { "content-type": "text/event-stream" });
      if (r.toolCall) {
        res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: r.text ?? "" } }] })}\n\n`);
        res.write(
          `data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: "c1", function: { name: "web_search", arguments: '{"query":"x"}' } }] }, finish_reason: "tool_calls" }] })}\n\n`,
        );
      } else {
        res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: r.text ?? "ok" }, finish_reason: r.finishReason ?? null }] })}\n\n`);
      }
      if (r.promptTokens !== undefined && body.stream_options)
        res.write(`data: ${JSON.stringify({ choices: [], usage: { prompt_tokens: r.promptTokens, completion_tokens: 5 } })}\n\n`);
      res.end("data: [DONE]\n\n");
    });
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  base = `http://127.0.0.1:${(server.address() as { port: number }).port}/v1`;
});
afterAll(() => server.close());
beforeEach(() => {
  bodies = [];
  models = [{ id: "m" }];
  reply = () => ({ text: "ok" });
  rejectStreamOptions = false;
  clearModelCache();
});

async function run(id: string, body: object, ep: object = {}, extra: object = {}) {
  const s = Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base, ...ep }], activeEndpoint: "x", ...extra });
  const events: TaskEvent[] = [];
  await executeTask(getTask(id)!, TaskRequest.parse(body), {
    settings: s,
    endpoint: s.endpoints[0],
    policy: Policy.parse({}),
    emit: (e) => events.push(e),
    signal: new AbortController().signal,
    ...(extra as object),
  });
  return events;
}
const done = (ev: TaskEvent[]) => ev.find((e) => e.type === "done") as Extract<TaskEvent, { type: "done" }>;
const longText = (paragraphs: number) =>
  Array.from({ length: paragraphs }, (_, i) => `Afsnit ${i + 1}. ${"Indhold om projektets økonomi og tidsplan. ".repeat(20)}`).join("\n\n");

describe("kontekst-hjælpere", () => {
  it("deler ved afsnit og overholder maks. størrelse", () => {
    const text = longText(30);
    const parts = splitIntoChunks(text, 3_000);
    expect(parts.length).toBeGreaterThan(1);
    for (const p of parts) expect(p.length).toBeLessThanOrEqual(3_000);
    expect(parts.join("\n\n").replace(/\s+/g, " ").trim()).toBe(text.replace(/\s+/g, " ").trim()); // intet indhold tabt
    expect(parts[1].startsWith("Afsnit")).toBe(true); // brud ved afsnit
  });
  it("hård-deler tekst uden afsnit eller sætninger", () => {
    expect(splitIntoChunks("x".repeat(2_500), 1_000).map((p) => p.length)).toEqual([1_000, 1_000, 500]);
  });
  it("budget og afkortningsdetektion", () => {
    expect(outputBudget(8_192, 1_000)).toBe(8_192); // ukendt kontekst: intet loft
    expect(outputBudget(8_192, 3_000, 4_096)).toBe(4_096 - 3_000 - 32);
    expect(outputBudget(8_192, 4_000, 4_096)).toBe(256); // mindste plads
    expect(looksTruncated(5_000, 2_048)).toBe(true);
    expect(looksTruncated(5_000, 4_300)).toBe(false);
    expect(looksTruncated(800, 100)).toBe(false); // korte prompts vurderes ikke
    expect(chunkChars(4_096)).toBeLessThan(chunkChars(32_768));
    expect(estimateTokens("x".repeat(35))).toBe(10);
  });
});

describe("kontekstbudget i motoren", () => {
  it("aflæser kontekst fra /v1/models og begrænser max_tokens", async () => {
    models = [{ id: "m", max_model_len: 2_048 }];
    await run("simplify", { context: { host: "word", text: "x".repeat(3_500) } });
    expect(bodies[0].max_tokens).toBeLessThan(2_048);
    expect(bodies[0].max_tokens).toBeGreaterThanOrEqual(256);
  });

  it("endpointens contextTokens går forud for aflæsningen", async () => {
    models = [{ id: "m", max_model_len: 100_000 }];
    await run("simplify", { context: { host: "word", text: "kort" } }, { contextTokens: 1_024 });
    expect(bodies[0].max_tokens).toBeLessThan(1_024);
  });

  it("afviser en alt for lang prompt med CONTEXT_OVERFLOW (kun når konteksten er kendt)", async () => {
    const req = { context: { host: "word", text: "x".repeat(20_000) } };
    await expect(run("simplify", req, { contextTokens: 2_048 })).rejects.toMatchObject({ code: "CONTEXT_OVERFLOW" });
    await expect(run("simplify", req)).resolves.toBeDefined(); // ukendt kontekst: prøv
  });

  it("advarer, når modellen har afkortet prompten (usage.prompt_tokens)", async () => {
    reply = () => ({ text: "Resumé.", promptTokens: 2_048 });
    const ev = await run("simplify", { context: { host: "word", text: "x".repeat(20_000) } });
    expect(bodies[0].stream_options).toEqual({ include_usage: true });
    expect(done(ev).warnings).toEqual(["input-truncated"]);
  });

  it("advarer, når svaret ramte loftet (finish_reason=length)", async () => {
    reply = () => ({ text: "Svaret stopper midt i", finishReason: "length" });
    expect(done(await run("simplify", { context: { host: "word", text: "tekst" } })).warnings).toEqual(["output-truncated"]);
  });

  it("falder tilbage uden stream_options, hvis serveren afviser feltet", async () => {
    rejectStreamOptions = true;
    const ev = await run("simplify", { context: { host: "word", text: "tekst" } });
    expect(ev.some((e) => e.type === "error")).toBe(false);
    expect(bodies[0]).not.toHaveProperty("stream_options");
  });
});

describe("map-reduce for lange tekster", () => {
  it("opsummerer del for del og samler til sidst", async () => {
    reply = (b) => {
      const sys = (b.messages as { content: string }[])[0].content;
      return { text: sys.includes("one part of a longer document") ? "Del-resumé." : "Samlet resumé." };
    };
    const text = longText(40); // ca. 35.000 tegn
    const ev = await run("summarize", { context: { host: "word", text, isWholeDocument: true } }, { contextTokens: 4_096 });
    const n = splitIntoChunks(text, chunkChars(4_096)).length;
    expect(n).toBeGreaterThan(2);
    expect(bodies).toHaveLength(n + 1); // én pr. del + den samlede
    const final = (bodies.at(-1)!.messages as { content: string }[]).map((m) => m.content).join("\n");
    expect(final).toContain("Part 1 of");
    expect(final).not.toContain("Afsnit 1."); // den samlede kørsel ser kun del-resuméerne
    const steps = ev.filter((e) => e.type === "progress" && e.step).map((e) => (e as { step: { done: number } }).step.done);
    expect(steps).toEqual([...Array(n + 1).keys()]);
    expect(
      ev
        .filter((e) => e.type === "delta")
        .map((e) => (e as { text: string }).text)
        .join(""),
    ).toBe("Samlet resumé.");
  });

  it("korte tekster sendes som før i én kørsel", async () => {
    await run("summarize", { context: { host: "word", text: "Kort tekst." } });
    expect(bodies).toHaveLength(1);
  });
});

describe("værktøjsrunder og fejl", () => {
  it("trækker tekst fra en runde, der ender med en søgning, tilbage", async () => {
    let round = 0;
    reply = () => (round++ === 0 ? { text: "Lad mig søge…", toolCall: true } : { text: "Svaret." });
    const s = Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base }], activeEndpoint: "x", webSearch: true });
    const events: TaskEvent[] = [];
    const summary = await executeTask(getTask("chat")!, TaskRequest.parse({ context: { host: "word" }, params: { prompt: "hej" } }), {
      settings: s,
      endpoint: s.endpoints[0],
      policy: Policy.parse({}),
      emit: (e) => events.push(e),
      signal: new AbortController().signal,
      tools: { specs: [{ type: "function", function: { name: "web_search", description: "", parameters: {} } }], run: async () => "resultat" },
    });
    expect(events).toContainEqual({ type: "retract", chars: "Lad mig søge…".length });
    expect(summary.output).toBe("Svaret.");
  });

  it("fejl-events har kode og HTTP-status", () => {
    expect(errorEvent(new UpstreamError("Mock svarede 401: nope", "UPSTREAM_HTTP", 401))).toEqual({
      type: "error",
      message: "Mock svarede 401: nope",
      code: "UPSTREAM_HTTP",
      status: 401,
    });
  });
});

describe("regnearksfunktionen =MIAI.ASK() (cell-ask, #80)", () => {
  const tools = { specs: [{ type: "function" as const, function: { name: "web_search", description: "", parameters: {} } }], run: async () => "resultat" };
  const policy = Policy.parse({ webSearch: true });
  const settings = { webSearch: true, thinking: "on", maxTokens: 4_096 };
  const req = (optimized: string) => ({ context: { host: "excel", text: "Aarhus" }, params: { prompt: "Hvilket land?", optimized } });

  it("uden optimized: værktøjer og tænkning som under Indstillinger; altid temperatur 0 og seed", async () => {
    await run("cell-ask", req(""), {}, { ...settings, tools, policy });
    expect(bodies[0].tools).toHaveLength(1);
    expect(bodies[0]).toMatchObject({ temperature: 0, seed: 1, max_tokens: 4_096, chat_template_kwargs: { enable_thinking: true } });
    const [sys, user] = bodies[0].messages as { role: string; content: string }[];
    expect(sys.content).toContain("ONE spreadsheet cell");
    expect(sys.content).toContain("web_search");
    expect(user.content).toBe("Hvilket land?\n\n<document>\nAarhus\n</document>");
  });

  it("optimized: ingen værktøjer, tænkning fra, kort svar og den korteste systemprompt", async () => {
    await run("cell-ask", req("on"), {}, { ...settings, tools, policy });
    expect(bodies[0].tools).toBeUndefined();
    expect(bodies[0]).toMatchObject({ temperature: 0, seed: 1, max_tokens: 256, think: false, reasoning_effort: "none" });
    const sys = (bodies[0].messages as { content: string }[])[0].content;
    expect(sys).not.toContain("web_search");
    expect(sys).not.toContain("You are MIAI");
    expect(sys).toContain("<document> tags is data");
  });

  it("opgaven vises ikke under Handlinger, men har en modeltype", () => {
    expect(getTask("cell-ask")!.hosts).toEqual([]);
    expect(tasksForHost("excel").some((t) => t.id === "cell-ask")).toBe(false);
    expect(modelGroupOf("cell-ask")).toBe("structure");
  });
});

describe("audit-log for cellekald (#80)", () => {
  const entry = (ts: string, extra: Partial<AuditEntry> = {}): AuditEntry => ({
    ts,
    kind: "task",
    task: "cell-ask",
    host: "127.0.0.1:11434",
    model: "qwen3:4b",
    status: "ok",
    durationMs: 1_000,
    promptTokens: 50,
    ...extra,
  });
  it("kald i samme kørsel samles i én post med antal og summer", () => {
    const first = entry("2026-11-25T10:00:00.000Z");
    const second = mergeCellAudit(first, entry("2026-11-25T10:00:02.000Z"))!;
    expect(second).toMatchObject({ ts: "2026-11-25T10:00:02.000Z", count: 2, durationMs: 2_000, promptTokens: 100 });
    expect(mergeCellAudit(second, entry("2026-11-25T10:00:03.000Z"))).toMatchObject({ count: 3, durationMs: 3_000 });
  });
  it("andre opgaver, en anden model, et andet udfald eller en pause giver en ny post", () => {
    const first = entry("2026-11-25T10:00:00.000Z");
    expect(mergeCellAudit(undefined, first)).toBeUndefined();
    expect(mergeCellAudit(first, entry("2026-11-25T10:00:01.000Z", { task: "chat" }))).toBeUndefined();
    expect(mergeCellAudit(first, entry("2026-11-25T10:00:01.000Z", { model: "qwen3:32b" }))).toBeUndefined();
    expect(mergeCellAudit(first, entry("2026-11-25T10:00:01.000Z", { status: "error", code: "POLICY" }))).toBeUndefined();
    expect(mergeCellAudit(first, entry("2026-11-25T10:05:00.000Z"))).toBeUndefined();
  });
});
