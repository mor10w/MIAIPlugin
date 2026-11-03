import { z } from "zod";
import { isDegenerating } from "./degenerate.js";
import { extractJson } from "./json.js";
import { listModels, streamChat, UpstreamError, type ClientOptions } from "./openai.js";
import { isUrlAllowed } from "./policy.js";
import type { TaskDefinition } from "./tasks.js";
import { ThinkFilter } from "./think.js";
import type { ChatMessage, Endpoint, ParsedTaskRequest, Policy, Settings, TaskEvent } from "./types.js";

/**
 * Opgavemotoren: bygger prompts, kalder endpointen, filtrerer <think>, validerer strukturerede svar.
 * Bruges både direkte i add-in'et og af bridgen (ADR-0007).
 */

export class InputError extends Error {}

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

export interface EngineDeps {
  settings: Settings;
  endpoint: Endpoint;
  policy: Policy;
  emit: (e: TaskEvent) => void;
  signal: AbortSignal;
  client?: ClientOptions;
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
  const guardLoops = task.output.kind === "text" && task.output.format !== "tsv";

  const run = async (msgs: ChatMessage[], responseFormat: Record<string, unknown> | undefined, onText: (t: string) => void) => {
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
    for await (const chunk of streamChat(ep, { model, messages: msgs, temperature, maxTokens: settings.maxTokens, responseFormat, signal }, client)) {
      if (chunk.text) push(chunk.text);
      if (chunk.usage) usage = chunk.usage;
      // Afbryd tidligt, hvis modellen er faldet i en gentagelsessløjfe. Brugeren får besked
      // via done.stopReason – der skrives intet ind i selve teksten, som kan ende i dokumentet.
      if (guardLoops && chunk.text && isDegenerating(full)) {
        stopReason = "repetition";
        break;
      }
    }
    if (filter) push(filter.flush());
    return full;
  };

  let output: string;
  if (task.output.kind === "text") {
    output = await run(messages, undefined, (text) => emit({ type: "delta", text }));
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
    const withSchema: ChatMessage[] = [
      { ...messages[0], content: `${messages[0].content}\n\nThe response must be a JSON object matching this JSON Schema:\n${JSON.stringify(jsonSchema)}` },
      ...messages.slice(1),
    ];
    let chars = 0;
    const progress = (t: string) => {
      chars += t.length;
      emit({ type: "progress", chars });
    };
    output = await run(withSchema, responseFormat, progress);
    let parsed = tryParse(schema, output);
    if (!parsed.ok) {
      // Ét genforsøg med fejlen som feedback – små modeller rammer ofte ved anden chance.
      const retry: ChatMessage[] = [
        ...withSchema,
        { role: "assistant", content: output },
        { role: "user", content: `That was not valid: ${parsed.error}. Reply again with only the JSON object.` },
      ];
      output = await run(retry, responseFormat, progress);
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
