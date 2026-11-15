import { createServer, type Server } from "node:http";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { clearModelCache, executeTask } from "./engine.js";
import { resetRejectedFields, thinkingFields } from "./openai.js";
import { getTask } from "./tasks.js";
import { Policy, Settings, TaskRequest, type TaskEvent } from "./types.js";

/** Tænkning og indsats (ADR-0003, tillæg): felterne i forespørgslen og tilbagefaldet hos strenge servere. */

let server: Server;
let base: string;
let bodies: Record<string, unknown>[] = [];
/** Afviser forespørgsler, der indeholder et af disse felter, med en OpenAI-lignende fejl. */
let rejects: string[] = [];
/** Fejlteksten ved en afvisning (standard: OpenAI's "Unrecognized request argument"). */
let rejectMessage: string | undefined;
/** Afviser alle forespørgsler med denne tekst (en fejl, der intet har med tænkning at gøre). */
let otherError: string | undefined;

beforeAll(async () => {
  server = createServer((req, res) => {
    if (req.url === "/v1/models") return res.end(JSON.stringify({ data: [{ id: "m" }] }));
    let raw = "";
    req.on("data", (c) => (raw += c));
    req.on("end", () => {
      const body = JSON.parse(raw) as Record<string, unknown>;
      bodies.push(body);
      const bad = rejects.find((f) => f in body);
      if (bad || otherError) {
        res.statusCode = 400;
        return res.end(JSON.stringify({ error: { message: otherError ?? rejectMessage ?? `Unrecognized request argument supplied: ${bad}` } }));
      }
      res.writeHead(200, { "content-type": "text/event-stream" });
      res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: "Kort svar." }, finish_reason: "stop" }] })}\n\n`);
      res.end("data: [DONE]\n\n");
    });
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  base = `http://127.0.0.1:${(server.address() as { port: number }).port}/v1`;
});
afterAll(() => server.close());
beforeEach(() => {
  bodies = [];
  rejects = [];
  rejectMessage = undefined;
  otherError = undefined;
  clearModelCache();
  resetRejectedFields();
});

async function run(settings: object = {}, ep: object = {}) {
  const s = Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: base, ...ep }], activeEndpoint: "x", ...settings });
  const events: TaskEvent[] = [];
  await executeTask(getTask("chat")!, TaskRequest.parse({ context: { host: "word" }, params: { prompt: "hej" } }), {
    settings: s,
    endpoint: s.endpoints[0],
    policy: Policy.parse({}),
    emit: (e) => events.push(e),
    signal: new AbortController().signal,
  });
  return events;
}
const done = (ev: TaskEvent[]) => ev.find((e) => e.type === "done") as Extract<TaskEvent, { type: "done" }> | undefined;
const FIELDS = ["reasoning_effort", "think", "chat_template_kwargs"];

describe("felterne for tænkning og indsats", () => {
  it("fra: reasoning_effort none, Ollamas think og chat-skabelonens enable_thinking", () => {
    expect(thinkingFields("off", undefined)).toEqual({
      standard: { reasoning_effort: "none" },
      extensions: { think: false, chat_template_kwargs: { enable_thinking: false } },
    });
    // Fra vinder over en indsats.
    expect(thinkingFields("off", "high").standard).toEqual({ reasoning_effort: "none" });
  });

  it("til: think og enable_thinking; med en indsats bærer reasoning_effort niveauet", () => {
    expect(thinkingFields("on", undefined)).toEqual({ standard: {}, extensions: { think: true, chat_template_kwargs: { enable_thinking: true } } });
    expect(thinkingFields("on", "high")).toEqual({ standard: { reasoning_effort: "high" }, extensions: { chat_template_kwargs: { enable_thinking: true } } });
  });

  it("serverens standard: kun en eventuel indsats, ellers intet", () => {
    expect(thinkingFields(undefined, "low")).toEqual({ standard: { reasoning_effort: "low" }, extensions: {} });
    expect(thinkingFields(undefined, undefined)).toEqual({ standard: {}, extensions: {} });
  });
});

describe("tænkning i forespørgslen", () => {
  it("standard sender ingen af felterne", async () => {
    await run();
    for (const f of FIELDS) expect(bodies[0]).not.toHaveProperty(f);
  });

  it("fra sendes i alle dialekter, og brugerens egne skabelon-argumenter bevares", async () => {
    await run({ thinking: "off" }, { extraBody: { chat_template_kwargs: { custom: 1 }, num_ctx: 8192 } });
    expect(bodies[0]).toMatchObject({ reasoning_effort: "none", think: false, chat_template_kwargs: { custom: 1, enable_thinking: false }, num_ctx: 8192 });
  });

  it("indsats uden tænkning til/fra sender kun reasoning_effort", async () => {
    await run({ reasoningEffort: "medium" });
    expect(bodies[0].reasoning_effort).toBe("medium");
    expect(bodies[0]).not.toHaveProperty("think");
    expect(bodies[0]).not.toHaveProperty("chat_template_kwargs");
  });
});

describe("tilbagefald hos servere, der afviser felterne", () => {
  it("afvises udvidelserne, sendes reasoning_effort alene – uden advarsel, og det huskes", async () => {
    rejects = ["think", "chat_template_kwargs"];
    const ev = await run({ thinking: "off" });
    expect(bodies).toHaveLength(2);
    expect(bodies[1]).toMatchObject({ reasoning_effort: "none" });
    expect(bodies[1]).not.toHaveProperty("think");
    expect(done(ev)?.warnings ?? []).not.toContain("thinking-ignored");

    bodies = [];
    await run({ thinking: "off" });
    expect(bodies).toHaveLength(1); // samme server og model: udvidelserne sendes ikke igen
    expect(bodies[0]).toMatchObject({ reasoning_effort: "none" });
  });

  it("afvises det hele, sendes forespørgslen uden, og svaret får en advarsel", async () => {
    rejects = FIELDS;
    const ev = await run({ thinking: "on", reasoningEffort: "high" });
    expect(bodies).toHaveLength(3);
    for (const f of FIELDS) expect(bodies[2]).not.toHaveProperty(f);
    expect(done(ev)?.warnings).toContain("thinking-ignored");
    expect(
      ev
        .filter((e) => e.type === "delta")
        .map((e) => (e as { text: string }).text)
        .join(""),
    ).toBe("Kort svar.");

    bodies = [];
    const again = await run({ thinking: "on", reasoningEffort: "high" });
    expect(bodies).toHaveLength(1);
    expect(done(again)?.warnings).toContain("thinking-ignored"); // advarslen gentages, så længe valget ikke virker
  });

  it("en model uden tænkning (Ollama: 'does not support thinking') får svaret uden tænkning og en advarsel", async () => {
    rejects = ["think"];
    rejectMessage = '"llama3.2:latest" does not support thinking';
    const ev = await run({ thinking: "on" });
    // Intet standardfelt at falde tilbage til: næste forsøg er helt uden tænke-felter.
    expect(bodies).toHaveLength(2);
    for (const f of FIELDS) expect(bodies[1]).not.toHaveProperty(f);
    expect(done(ev)?.warnings).toContain("thinking-ignored");
  });

  it("andre 400-fejl prøves ikke igen og glemmer ikke tænkningen", async () => {
    otherError = "maximum context length exceeded";
    await expect(run({ thinking: "off" })).rejects.toThrow(/context length/);
    expect(bodies).toHaveLength(1);
    otherError = undefined;
    bodies = [];
    await run({ thinking: "off" });
    expect(bodies[0]).toMatchObject({ think: false });
  });
});
