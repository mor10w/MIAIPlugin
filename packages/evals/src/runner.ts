import { clearModelCache, executeTask, Policy, Settings, TaskRequest, TASKS, type TaskEvent, type ToolExecutor } from "@miai/shared";
import { runCheck } from "./checks.js";
import type { CaseResult, EvalCase } from "./types.js";

/**
 * Kører eval-cases gennem produktets egen motor (`executeTask` fra @miai/shared), så præcis de
 * prompts, den validering og de filtre, som add-in'et og bridgen bruger, bliver evalueret (#2).
 */

export interface EndpointConfig {
  baseUrl: string;
  apiKey?: string;
  structuredOutput: "json_schema" | "json_object" | "none";
  presencePenalty?: number;
  timeoutMs?: number;
}

export interface JudgeConfig {
  baseUrl: string;
  apiKey?: string;
  model: string;
}

/** Konfigurationens navn i rapporten, fx "json_schema/pp1.5". */
export const configName = (c: EndpointConfig) => `${c.structuredOutput}/pp${c.presencePenalty ?? 0}`;

/** Faste søgeresultater: web_search-evalueringen måler, om modellen kalder værktøjet – intet sendes på nettet. */
const FAKE_RESULTS = (query: string) =>
  `Results for "${query}":\n1. Example result – https://example.org/1\n   Snippet: Information related to ${query}.\n2. Another source – https://example.org/2\n   Snippet: More details about ${query}.`;

function fakeTools(): ToolExecutor {
  return {
    specs: [
      {
        type: "function",
        function: {
          name: "web_search",
          description:
            "Search the web for current information, news, prices, versions or facts you are not sure about. Returns a few results with title, URL and snippet.",
          parameters: { type: "object", properties: { query: { type: "string", description: "The search query" } }, required: ["query"] },
        },
      },
    ],
    run: async (_name, args) => FAKE_RESULTS((JSON.parse(args || "{}") as { query?: string }).query ?? ""),
  };
}

/** Tæller kald til /chat/completions, så "gyldigt i første forsøg" kan måles uden at ændre motoren. */
function countingFetch(counter: { n: number }): typeof fetch {
  const orig = globalThis.fetch;
  return (input, init) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    if (url.endsWith("/chat/completions")) counter.n++;
    return orig(input, init);
  };
}

export async function runCase(c: EvalCase, model: string, ep: EndpointConfig, judge?: JudgeConfig): Promise<CaseResult> {
  const task = TASKS.find((t) => t.id === c.task);
  if (!task) throw new Error(`${c.id}: ukendt opgave '${c.task}'`);
  const structured = task.output.kind === "json";
  const settings = Settings.parse({
    endpoints: [
      {
        id: "eval",
        name: "Eval",
        baseUrl: ep.baseUrl,
        apiKey: ep.apiKey ?? "",
        structuredOutput: ep.structuredOutput,
        presencePenalty: ep.presencePenalty || undefined,
        timeoutMs: ep.timeoutMs ?? 300_000,
      },
    ],
    activeEndpoint: "eval",
    defaultModel: model,
    webSearch: c.tools,
  });
  const req = TaskRequest.parse({ context: c.context, params: c.params, history: c.history, model });

  let output = "";
  let result: unknown;
  let repetition = false;
  const warnings: string[] = [];
  const toolCalls: string[] = [];
  let completionTokens: number | undefined;
  const emit = (e: TaskEvent) => {
    if (e.type === "delta") output += e.text;
    else if (e.type === "retract") output = output.slice(0, Math.max(0, output.length - e.chars));
    else if (e.type === "result") result = e.data;
    else if (e.type === "search") toolCalls.push(e.query);
    else if (e.type === "done") {
      repetition = e.stopReason === "repetition";
      warnings.push(...(e.warnings ?? []));
      completionTokens = e.usage?.completionTokens;
    }
  };

  const counter = { n: 0 };
  const realFetch = globalThis.fetch;
  globalThis.fetch = countingFetch(counter);
  const started = Date.now();
  let error: string | undefined;
  try {
    const summary = await executeTask(task, req, {
      settings,
      endpoint: settings.endpoints[0],
      policy: Policy.parse({ allowedHosts: ["*"] }),
      emit,
      signal: AbortSignal.timeout(ep.timeoutMs ?? 300_000),
      tools: c.tools ? fakeTools() : undefined,
    });
    if (structured) output = summary.output;
    else output = summary.output || output;
  } catch (err) {
    error = (err as Error).message;
  } finally {
    globalThis.fetch = realFetch;
  }
  const durationMs = Date.now() - started;

  const input = [c.context.text, c.context.table, c.params.prompt].filter(Boolean).join("\n");
  const checks = error ? [] : c.checks.map((ch) => runCheck(ch, { output, result, input, toolCalls }));
  const res: CaseResult = {
    caseId: c.id,
    task: c.task,
    lang: c.lang,
    model,
    config: configName(ep),
    ok: !error && checks.every((x) => x.ok),
    ...(error ? { error } : {}),
    output,
    ...(result !== undefined ? { result } : {}),
    checks,
    requests: counter.n,
    structured,
    ...(structured ? { firstTryValid: !error && counter.n === 1, validAfterRetry: !error } : {}),
    repetition,
    toolCalls,
    warnings,
    durationMs,
    completionTokens,
  };
  if (judge && c.rubric && !error) res.judge = await judgeAnswer(judge, c, output);
  return res;
}

/** Dommermodel: bedømmer svaret 1–5 efter casens kriterier. */
export async function judgeAnswer(j: JudgeConfig, c: EvalCase, output: string): Promise<{ score: number; reason: string }> {
  const input = [c.context.text, c.context.table, c.params.prompt && `Instruction: ${c.params.prompt}`].filter(Boolean).join("\n\n");
  const body = {
    model: j.model,
    temperature: 0,
    response_format: { type: "json_object" },
    messages: [
      {
        role: "system",
        content:
          'You grade answers from an Office writing assistant. Score 1 (unusable) to 5 (excellent) against the criteria. Be strict about factual errors, wrong language and invented content. Reply with JSON: {"score": <1-5>, "reason": "<one sentence>"}.',
      },
      { role: "user", content: `Task: ${c.task}\nCriteria: ${c.rubric}\n\n<input>\n${input}\n</input>\n\n<answer>\n${output}\n</answer>` },
    ],
  };
  try {
    const r = await fetch(`${j.baseUrl.replace(/\/$/, "")}/chat/completions`, {
      method: "POST",
      headers: { "content-type": "application/json", ...(j.apiKey ? { authorization: `Bearer ${j.apiKey}` } : {}) },
      body: JSON.stringify(body),
    });
    const data = (await r.json()) as { choices?: { message?: { content?: string } }[] };
    const text = data.choices?.[0]?.message?.content ?? "";
    const parsed = JSON.parse(text.slice(text.indexOf("{"), text.lastIndexOf("}") + 1)) as { score?: number; reason?: string };
    const score = Math.min(5, Math.max(1, Math.round(Number(parsed.score))));
    return { score: Number.isFinite(score) ? score : 0, reason: parsed.reason ?? "" };
  } catch (err) {
    return { score: 0, reason: `Dommeren fejlede: ${(err as Error).message}` };
  }
}

export interface RunOptions {
  models: string[];
  configs: EndpointConfig[];
  judge?: JudgeConfig;
  repeat?: number;
  onResult?: (r: CaseResult, i: number, total: number) => void;
}

/** Kører alle cases for hver model × konfiguration × gentagelse – sekventielt (én lokal GPU). */
export async function runMatrix(cases: EvalCase[], o: RunOptions): Promise<CaseResult[]> {
  const out: CaseResult[] = [];
  const total = cases.length * o.models.length * o.configs.length * (o.repeat ?? 1);
  for (const model of o.models)
    for (const cfg of o.configs) {
      clearModelCache();
      for (let rep = 0; rep < (o.repeat ?? 1); rep++)
        for (const c of cases) {
          const r = await runCase(c, model, cfg, o.judge);
          out.push(r);
          o.onResult?.(r, out.length, total);
        }
    }
  return out;
}
