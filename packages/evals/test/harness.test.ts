import { createServer, type IncomingMessage, type Server } from "node:http";
import type { AddressInfo } from "node:net";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { runCheck, detectLanguage } from "../src/checks.js";
import { loadCases } from "../src/cli.js";
import { aggregate, bestPerCategory, markdownReport } from "../src/report.js";
import { runCase } from "../src/runner.js";
import { EvalCase, type CaseResult, type EvalCaseInput } from "../src/types.js";
import { TASKS } from "@miai/shared";

/** Harnesset (#2) testet mod en mock OpenAI-server – ingen rigtig model nødvendig. */

type Reply = { text: string } | { tool: string };
let replies: Reply[] = [];
const bodies: { stream?: boolean; tools?: unknown[] }[] = [];
let server: Server;
let baseUrl: string;

const readBody = (req: IncomingMessage) =>
  new Promise<string>((ok) => {
    let s = "";
    req.on("data", (c) => (s += c));
    req.on("end", () => ok(s));
  });

beforeAll(async () => {
  server = createServer(async (req, res) => {
    if (req.url?.endsWith("/models")) return res.end(JSON.stringify({ object: "list", data: [{ id: "mock" }] }));
    const body = JSON.parse(await readBody(req));
    bodies.push(body);
    const r = replies.shift() ?? { text: "" };
    if (!body.stream) {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ choices: [{ message: { content: "text" in r ? r.text : "" }, finish_reason: "stop" }] }));
    }
    res.writeHead(200, { "content-type": "text/event-stream" });
    const chunk =
      "tool" in r
        ? {
            choices: [
              {
                delta: {
                  tool_calls: [{ index: 0, id: "c1", type: "function", function: { name: "web_search", arguments: JSON.stringify({ query: r.tool }) } }],
                },
                finish_reason: "tool_calls",
              },
            ],
          }
        : { choices: [{ delta: { content: r.text }, finish_reason: "stop" }] };
    res.write(`data: ${JSON.stringify(chunk)}\n\n`);
    res.end("data: [DONE]\n\n");
  });
  await new Promise<void>((ok) => server.listen(0, "127.0.0.1", ok));
  baseUrl = `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`;
});
afterAll(() => new Promise<void>((ok) => server.close(() => ok())));

const ep = () => ({ baseUrl, structuredOutput: "json_schema" as const });
const mk = (c: Partial<EvalCaseInput> & Pick<EvalCaseInput, "id" | "task">) =>
  EvalCase.parse({ lang: "da", context: { host: "word", text: "Hej verden." }, ...c });

describe("eval-sættet", () => {
  const cases = loadCases();
  it("har 30–60 cases på både dansk og engelsk (#2: 30–50, udvidet med agent og design)", () => {
    expect(cases.length).toBeGreaterThanOrEqual(30);
    expect(cases.length).toBeLessThanOrEqual(60);
    expect(cases.filter((c) => c.lang === "da").length).toBeGreaterThan(10);
    expect(cases.filter((c) => c.lang === "en").length).toBeGreaterThan(10);
  });
  it("bruger kun eksisterende opgaver, og alle indbyggede opgaver er dækket", () => {
    const ids = new Set(TASKS.map((t) => t.id));
    for (const c of cases) expect(ids.has(c.task), c.id).toBe(true);
    for (const id of ids)
      expect(
        cases.some((c) => c.task === id),
        id,
      ).toBe(true);
  });
  it("hver case har mindst ét automatisk tjek og en rubric", () => {
    for (const c of cases) {
      expect(c.checks.length, c.id).toBeGreaterThan(0);
      expect(c.rubric, c.id).toBeTruthy();
    }
  });
});

describe("tjek", () => {
  const a = (output: string, extra: Partial<Parameters<typeof runCheck>[1]> = {}) => ({ output, input: "et to tre fire fem", toolCalls: [], ...extra });
  it("sprog", () => {
    expect(detectLanguage("Mødet er flyttet til torsdag, og det er ikke muligt at komme")).toBe("da");
    expect(detectLanguage("The meeting has been moved to Thursday and it is not possible")).toBe("en");
    expect(detectLanguage("Das Treffen ist auf Donnerstag verschoben und wir sind nicht da")).toBe("de");
  });
  it("forord og kodeblokke", () => {
    expect(runCheck({ type: "noPreamble" }, a("Her er den omskrevne tekst: ...")).ok).toBe(false);
    expect(runCheck({ type: "noPreamble" }, a("```\ntekst\n```")).ok).toBe(false);
    expect(runCheck({ type: "noPreamble" }, a("Mødet er flyttet.")).ok).toBe(true);
  });
  it("længde i forhold til input", () => {
    expect(runCheck({ type: "length", maxRatio: 1 }, a("et to tre fire fem seks")).ok).toBe(false);
    expect(runCheck({ type: "length", maxRatio: 1.5 }, a("et to tre fire fem seks")).ok).toBe(true);
  });
  it("punkter, TSV og regex", () => {
    expect(runCheck({ type: "bullets", min: 3 }, a("- a\n- b\n* c\n1. d")).ok).toBe(true);
    expect(runCheck({ type: "tsv", columns: 2, rows: 2 }, a("a\tb\nc\td")).ok).toBe(true);
    expect(runCheck({ type: "tsv", columns: 2 }, a("a\tb\nc")).ok).toBe(false);
    expect(runCheck({ type: "regex", pattern: "^banana$", negate: true }, a("BANANA")).ok).toBe(false);
  });
  it("strukturerede svar og grammatik", () => {
    const result = { formula: '=SUMIF(A:A,"Jylland",B:B)', edits: [{ original: "bestiling", replacement: "bestilling", explanation: "" }] };
    expect(runCheck({ type: "result", path: "formula", pattern: "^=SUMIFS?\\(" }, a("", { result })).ok).toBe(true);
    expect(runCheck({ type: "edits", expect: [{ fix: "bestilling" }], maxEdits: 1 }, a("", { result })).ok).toBe(true);
    expect(runCheck({ type: "edits", expect: [], maxEdits: 0 }, a("", { result })).ok).toBe(false);
  });
});

describe("runner", () => {
  it("fritekst gennem produktets motor: tjek og <think> filtreret", async () => {
    replies = [{ text: "<think>overvejer</think>Mødet er flyttet til torsdag, og det er ikke et problem." }];
    const r = await runCase(
      mk({
        id: "t1",
        task: "rewrite",
        checks: [
          { type: "language", lang: "da" },
          { type: "notContains", values: ["overvejer"] },
        ],
      }),
      "mock",
      ep(),
    );
    expect(r.error).toBeUndefined();
    expect(r.ok).toBe(true);
    expect(r.requests).toBe(1);
    expect(r.structured).toBe(false);
  });

  it("struktureret svar: gyldigt i første forsøg", async () => {
    replies = [{ text: JSON.stringify({ formula: "=SUM(B:B)", explanation: "Summen" }) }];
    const r = await runCase(
      mk({
        id: "t2",
        task: "generate-formula",
        context: { host: "excel", text: "" },
        params: { prompt: "sum af B" },
        checks: [{ type: "result", path: "formula", pattern: "^=SUM" }],
      }),
      "mock",
      ep(),
    );
    expect(r).toMatchObject({ ok: true, structured: true, firstTryValid: true, validAfterRetry: true, requests: 1 });
  });

  it("struktureret svar: gyldigt først efter genforsøg tælles som sådan", async () => {
    replies = [{ text: "Her er formlen: =SUM(B:B)" }, { text: JSON.stringify({ formula: "=SUM(B:B)", explanation: "Summen" }) }];
    const r = await runCase(mk({ id: "t3", task: "generate-formula", context: { host: "excel", text: "" }, params: { prompt: "sum af B" } }), "mock", ep());
    expect(r).toMatchObject({ ok: true, firstTryValid: false, validAfterRetry: true, requests: 2 });
  });

  it("ugyldigt svar efter genforsøg giver fejl, ikke et crash", async () => {
    replies = [{ text: "nej" }, { text: "stadig nej" }];
    const r = await runCase(mk({ id: "t4", task: "generate-formula", context: { host: "excel", text: "" }, params: { prompt: "x" } }), "mock", ep());
    expect(r.ok).toBe(false);
    expect(r.validAfterRetry).toBe(false);
    expect(r.error).toMatch(/gyldigt struktureret svar/);
  });

  it("web_search: værktøjet tilbydes kun med tools, og kald registreres uden netværk", async () => {
    replies = [{ tool: "node.js latest version" }, { text: "Den seneste version er 24 (kilde: https://example.org/1)." }];
    const r = await runCase(
      mk({ id: "t5", task: "chat", params: { prompt: "Seneste Node?" }, tools: true, checks: [{ type: "tool", called: true }] }),
      "mock",
      ep(),
    );
    expect(r.ok).toBe(true);
    expect(r.toolCalls).toEqual(["node.js latest version"]);
    expect(bodies.at(-2)?.tools).toHaveLength(1);

    replies = [{ text: "Svar." }];
    await runCase(mk({ id: "t6", task: "chat", params: { prompt: "Hej" } }), "mock", ep());
    expect(bodies.at(-1)?.tools).toBeUndefined();
  });

  it("gentagelsessløjfer registreres", async () => {
    replies = [{ text: "Start. " + "side 6 tabel ".repeat(40) }];
    const r = await runCase(mk({ id: "t7", task: "rewrite" }), "mock", ep());
    expect(r.repetition).toBe(true);
  });
});

describe("rapport", () => {
  const base: Omit<CaseResult, "caseId" | "task" | "model" | "ok"> = {
    lang: "da",
    config: "json_schema/pp0",
    output: "",
    checks: [],
    requests: 1,
    structured: false,
    repetition: false,
    toolCalls: [],
    warnings: [],
    durationMs: 1000,
  };
  const rs: CaseResult[] = [
    { ...base, caseId: "a", task: "summarize", model: "lille", ok: false },
    { ...base, caseId: "a", task: "summarize", model: "stor", ok: true },
    { ...base, caseId: "f", task: "generate-formula", model: "lille", ok: true, structured: true, firstTryValid: true, validAfterRetry: true },
    {
      ...base,
      caseId: "f",
      task: "generate-formula",
      model: "stor",
      ok: true,
      structured: true,
      firstTryValid: false,
      validAfterRetry: true,
      durationMs: 3000,
    },
  ];
  it("aggregerer og vælger bedste model pr. kategori", () => {
    expect(aggregate(rs.filter((r) => r.model === "stor"))).toMatchObject({ n: 2, passRate: 1, firstTry: 0, afterRetry: 1 });
    const best = Object.fromEntries(bestPerCategory(rs).map((b) => [b.category, b.model]));
    expect(best).toEqual({ "Resumé og analyse": "stor", Formler: "lille" });
  });
  it("skriver en Markdown-rapport med fejlede cases", () => {
    const md = markdownReport(rs, { endpoint: "http://x/v1", date: "2026-11-03" });
    expect(md).toContain("## Bedste model pr. kategori");
    expect(md).toContain("| a | lille |");
  });
});

describe("agent-planer (#6)", () => {
  const doc = "Vi har mange kunder. Alle kunder får svar.";
  const ans = (actions: unknown[]) => ({ output: "", result: { summary: "", actions }, input: doc, toolCalls: [], host: "word" as const });
  it("kræver de rigtige handlinger, forbyder destruktive og tjekker, at teksten findes", () => {
    const good = ans([{ action: "replace", find: "kunder", replace: "klienter", all: true, reason: "" }]);
    expect(runCheck({ type: "plan", min: 1, include: ["replace"], exclude: ["delete"], valid: true }, good).ok).toBe(true);
    const bad = ans([{ action: "delete", find: "findes ikke", reason: "" }]);
    const r = runCheck({ type: "plan", include: ["replace"], exclude: ["delete"], valid: true }, bad);
    expect(r.ok).toBe(false);
    expect(r.detail).toMatch(/mangler replace.*indeholder delete.*kan ikke udføres/);
  });
});
