import { z } from "zod";
import { isDegenerating } from "./degenerate.js";
import { extractJson } from "./json.js";
import { listModels, streamChat, UpstreamError, type ApiMessage, type ClientOptions, type ToolCall, type ToolSpec } from "./openai.js";
import { isUrlAllowed } from "./policy.js";
import { textFormat, type TaskDefinition } from "./tasks.js";
import { ThinkFilter } from "./think.js";
import type { Endpoint, ParsedTaskRequest, Policy, Settings, TaskEvent } from "./types.js";

/**
 * Opgavemotoren: bygger prompts, kalder endpointen, filtrerer <think>, validerer strukturerede svar.
 * Bruges både direkte i add-in'et og af bridgen (ADR-0007).
 */

export class InputError extends Error {}

/** Maks. antal værktøjsrunder pr. opgave – beskytter mod uendelige søge-løver. */
const MAX_TOOL_ROUNDS = 4;

/** Linje der tilføjes til systemprompten, når modellen kan søge (ADR-0010). */
const SEARCH_AWARENESS =
  "You have access to a web_search tool that searches the web. Use it when you need current information (news, prices, versions, recent events) or a fact you are not confident about; do not use it when the document already answers the question. If your answer uses web results, cite sources with their URLs.";

/** Sidste runde (maks. antal søgninger nået): modellen skal svare med det, den har fundet. */
const FINAL_ANSWER_NUDGE = "You have used all your searches. Stop searching now and give the final answer to the user based on what you found.";

export function validateInput(task: TaskDefinition, req: ParsedTaskRequest): void {
  const c = req.context;
  const hasText = c.text.trim().length > 0 || !!c.table?.trim();
  switch (task.input) {
    case "selection":
    case "selectionOrDocument":
      if (!hasText) throw new InputError("Markér noget tekst først.");
      break;
    case "formula":
      if (!c.formula?.startsWith("=")) throw new InputError("Markér en celle, der indeholder en formel.");
      break;
    case "prompt":
      if (!req.params.prompt?.trim() && !hasText) throw new InputError("Skriv en instruktion eller markér tekst.");
      break;
  }
}

const modelCache = new Map<string, { models: string[]; at: number }>();

export async function cachedModels(ep: Endpoint, signal?: AbortSignal, opts?: ClientOptions): Promise<string[]> {
  const key = `${ep.baseUrl}|${ep.apiKey}`;
  const hit = modelCache.get(key);
  if (hit && Date.now() - hit.at < 30_000) return hit.models;
  const models = await listModels(ep, signal, opts);
  modelCache.set(key, { models, at: Date.now() });
  return models;
}

export function clearModelCache(): void {
  modelCache.clear();
}

/** Værktøjer, modellen kan kalde (ADR-0010). Bridgen leverer eksekveringen; browseren gør det ikke. */
export interface ToolExecutor {
  specs: ToolSpec[];
  /** Kører et værktøjskald og returnerer resultatet som tekst til modellen. */
  run(name: string, argsJson: string): Promise<string>;
}

export interface EngineDeps {
  settings: Settings;
  endpoint: Endpoint;
  policy: Policy;
  emit: (e: TaskEvent) => void;
  signal: AbortSignal;
  client?: ClientOptions;
  /** Leveres kun af bridgen – direkte tilstand har ingen værktøjer. */
  tools?: ToolExecutor;
}

export interface TaskSummary {
  model: string;
  output: string;
  durationMs: number;
  inputChars: number;
  usage?: { promptTokens?: number; completionTokens?: number };
}

/** Kører en opgave og udsender TaskEvents. Kaster ved fejl (kalderen omsætter til en "error"-event). */
export async function executeTask(task: TaskDefinition, req: ParsedTaskRequest, deps: EngineDeps): Promise<TaskSummary> {
  const { settings, endpoint: ep, policy, emit, signal, client } = deps;
  if (!isUrlAllowed(ep.baseUrl, policy)) {
    throw new UpstreamError(`Endpoint ${ep.baseUrl} er ikke tilladt af politikken.`, "POLICY");
  }

  const model = req.model || settings.taskModels[task.id] || settings.defaultModel || (await cachedModels(ep, signal, client))[0];
  if (!model) throw new UpstreamError(`${ep.name} har ingen modeller. Hent eller indlæs en model i dit inference-værktøj.`, "NO_MODEL");
  emit({ type: "meta", model, endpoint: ep.name });

  const messages = task.buildMessages(req);
  const temperature = task.temperature ?? settings.temperature;
  const started = Date.now();
  let usage: TaskSummary["usage"];
  let stopReason: "repetition" | undefined;
  // Sløjfe-vagten gælder kun fritekst. TSV (Excel-celler) og JSON har legitime gentagelser
  // og valideres på anden vis.
  const format = textFormat(task, req.context);
  const guardLoops = format === "plain" || format === "markdown";
  // presence_penalty følger samme afgrænsning: den straffer gentagne tokens, hvilket hjælper
  // mod sløjfer i fritekst, men skader JSON (gentagne nøgler) og tabeller (gentagne værdier).
  const presencePenalty = guardLoops && ep.presencePenalty ? ep.presencePenalty : undefined;

  // Værktøjer (ADR-0010): kun når bridgen leverer en eksekverer, brugeren har slået web-søgning
  // til, og politikken tillader det (søgeord sendes til en ekstern tjeneste).
  const tools = deps.tools && settings.webSearch && policy.webSearch ? deps.tools : undefined;

  const addUsage = (u: { promptTokens?: number; completionTokens?: number } | undefined) => {
    if (!u) return;
    usage = {
      promptTokens: (usage?.promptTokens ?? 0) + (u.promptTokens ?? 0),
      completionTokens: (usage?.completionTokens ?? 0) + (u.completionTokens ?? 0),
    };
  };

  const run = async (msgs: ApiMessage[], responseFormat: Record<string, unknown> | undefined, onText: (t: string) => void, useTools = true): Promise<{ text: string; toolCalls?: ToolCall[] }> => {
    const filter = ep.stripThinking ? new ThinkFilter() : undefined;
    let full = "";
    let begun = false;
    const push = (raw: string) => {
      let t = filter ? filter.push(raw) : raw;
      if (!begun) {
        t = t.replace(/^\s+/, "");
        if (!t) return;
        begun = true;
      }
      if (!t) return;
      full += t;
      onText(t);
    };
    // Værktøjskald bygges op over stream-chunks (arguments ankommer som fragmenter).
    const tcParts = new Map<number, { id: string; name: string; args: string }>();
    for await (const chunk of streamChat(ep, { model, messages: msgs, temperature, maxTokens: settings.maxTokens, presencePenalty, responseFormat, signal, tools: useTools ? tools?.specs : undefined }, client)) {
      if (chunk.text) push(chunk.text);
      addUsage(chunk.usage);
      for (const d of chunk.toolCalls ?? []) {
        const p = tcParts.get(d.index) ?? { id: "", name: "", args: "" };
        if (d.id) p.id = d.id;
        if (d.function?.name) p.name += d.function.name;
        if (d.function?.arguments) p.args += d.function.arguments;
        tcParts.set(d.index, p);
      }
      // Afbryd tidligt, hvis modellen er faldet i en gentagelsessløjfe. Brugeren får besked
      // via done.stopReason – der skrives intet ind i selve teksten, som kan ende i dokumentet.
      if (guardLoops && chunk.text && isDegenerating(full)) {
        stopReason = "repetition";
        break;
      }
    }
    if (filter) push(filter.flush());
    const toolCalls = tcParts.size
      ? [...tcParts.entries()].sort((a, b) => a[0] - b[0]).map(([i, p]) => ({ id: p.id || `call_${i}`, function: { name: p.name, arguments: p.args } }))
      : undefined;
    return { text: full, toolCalls };
  };

  let output: string;
  if (task.output.kind === "text") {
    // Modellen skal være klar over, den kan søge (ADR-0010).
    let msgs: ApiMessage[] = messages;
    if (tools) {
      const [sys, ...rest] = msgs;
      msgs = [{ ...sys, content: `${sys.content}\n${SEARCH_AWARENESS}` }, ...rest];
    }
    // Løkke: modellen kan kalde web_search, hvorefter resultatet sendes tilbage som tool-besked.
    // Ved MAX_TOOL_ROUNDS tillades ingen nye kald – modellen skal svare med det, den har.
    for (let round = 0; ; round++) {
      const lastRound = round >= MAX_TOOL_ROUNDS;
      // Sidste runde: ingen værktøjer, og en opfordring til at svare med det fundne.
      const roundMsgs = lastRound ? [...msgs, { role: "user" as const, content: FINAL_ANSWER_NUDGE }] : msgs;
      const r = await run(roundMsgs, undefined, (text) => emit({ type: "delta", text }), !lastRound);
      if (!r.toolCalls?.length || lastRound || !tools) {
        output = r.text;
        break;
      }
      const toolMsgs: ApiMessage[] = [{ role: "assistant", content: r.text, tool_calls: r.toolCalls.map((tc) => ({ ...tc, type: "function" as const })) }];
      for (const tc of r.toolCalls) {
        if (tc.function.name === "web_search") {
          let query = "";
          try {
            query = ((JSON.parse(tc.function.arguments) as { query?: string }).query ?? "").slice(0, 200);
          } catch {
            // Ugyldige argumenter: vis dem som de er.
          }
          emit({ type: "search", query });
        }
        const result = await tools.run(tc.function.name, tc.function.arguments).catch((err) => `Værktøjet fejlede: ${(err as Error).message}`);
        toolMsgs.push({ role: "tool", tool_call_id: tc.id, content: result });
      }
      msgs = [...msgs, ...toolMsgs];
    }
  } else {
    const schema = task.output.schema;
    const jsonSchema = z.toJSONSchema(schema) as Record<string, unknown>;
    const responseFormat =
      ep.structuredOutput === "json_schema"
        ? { type: "json_schema", json_schema: { name: task.output.schemaName, schema: jsonSchema } }
        : ep.structuredOutput === "json_object"
          ? { type: "json_object" }
          : undefined;
    // Skemaet står også i prompten – hjælper små modeller og endpoints uden struktureret output.
    const withSchema: ApiMessage[] = [
      { ...messages[0], content: `${messages[0].content}\n\nThe response must be a JSON object matching this JSON Schema:\n${JSON.stringify(jsonSchema)}` },
      ...messages.slice(1),
    ];
    let chars = 0;
    const progress = (t: string) => {
      chars += t.length;
      emit({ type: "progress", chars });
    };
    output = (await run(withSchema, responseFormat, progress, false)).text;
    let parsed = tryParse(schema, output);
    if (!parsed.ok) {
      // Ét genforsøg med fejlen som feedback – små modeller rammer ofte ved anden chance.
      const retry: ApiMessage[] = [
        ...withSchema,
        { role: "assistant", content: output },
        { role: "user", content: `That was not valid: ${parsed.error}. Reply again with only the JSON object.` },
      ];
      output = (await run(retry, responseFormat, progress, false)).text;
      parsed = tryParse(schema, output);
    }
    if (!parsed.ok) throw new UpstreamError(`Modellen returnerede ikke et gyldigt struktureret svar (${parsed.error}). Prøv en større model.`, "UPSTREAM_PROTOCOL");
    emit({ type: "result", data: parsed.data });
  }

  emit({ type: "done", usage, stopReason });
  return { model, output, usage, durationMs: Date.now() - started, inputChars: messages.reduce((n, m) => n + m.content.length, 0) };
}

/** Omsætter en fejl til en TaskEvent. */
export function errorEvent(err: unknown): TaskEvent {
  const code = err instanceof UpstreamError ? err.code : err instanceof InputError ? "NO_INPUT" : "INTERNAL";
  return { type: "error", message: (err as Error)?.message ?? String(err), code };
}

function tryParse(schema: z.ZodType, text: string): { ok: true; data: unknown } | { ok: false; error: string } {
  try {
    const r = schema.safeParse(extractJson(text));
    return r.success ? { ok: true, data: r.data } : { ok: false, error: z.prettifyError(r.error).slice(0, 500) };
  } catch (err) {
    return { ok: false, error: (err as Error).message };
  }
}
