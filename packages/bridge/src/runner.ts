import { appendFile } from "node:fs/promises";
import { z } from "zod";
import type { ChatMessage, Endpoint, ParsedTaskRequest, TaskDefinition, TaskEvent } from "@miai/shared";
import type { ConfigStore } from "./config.js";
import { extractJson } from "./json.js";
import { listModels, streamChat, UpstreamError } from "./openai.js";
import { isUrlAllowed } from "./policy.js";
import { ThinkFilter } from "./think.js";

export class InputError extends Error {}

const modelCache = new Map<string, { models: string[]; at: number }>();

export async function cachedModels(ep: Endpoint, signal?: AbortSignal): Promise<string[]> {
  const key = `${ep.baseUrl}|${ep.apiKey}`;
  const hit = modelCache.get(key);
  if (hit && Date.now() - hit.at < 30_000) return hit.models;
  const models = await listModels(ep, signal);
  modelCache.set(key, { models, at: Date.now() });
  return models;
}

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

interface RunOptions {
  task: TaskDefinition;
  req: ParsedTaskRequest;
  store: ConfigStore;
  emit: (e: TaskEvent) => void;
  signal: AbortSignal;
}

export async function runTask({ task, req, store, emit, signal }: RunOptions): Promise<void> {
  const settings = store.getSettings();
  const ep = store.activeEndpoint();
  if (!isUrlAllowed(ep.baseUrl, store.getPolicy())) {
    throw new UpstreamError(`Endpoint ${ep.baseUrl} er ikke tilladt af politikken.`, "UPSTREAM_UNREACHABLE");
  }

  const model = req.model || settings.taskModels[task.id] || settings.defaultModel || (await cachedModels(ep, signal))[0];
  if (!model) throw new UpstreamError(`${ep.name} har ingen modeller. Hent eller indlæs en model i dit inference-værktøj.`, "NO_MODEL");
  emit({ type: "meta", model, endpoint: ep.name });

  const messages = task.buildMessages(req);
  const temperature = task.temperature ?? settings.temperature;
  const started = Date.now();
  let usage: { promptTokens?: number; completionTokens?: number } | undefined;

  const run = async (msgs: ChatMessage[], responseFormat: Record<string, unknown> | undefined, onText: (t: string) => void) => {
    const filter = ep.stripThinking ? new ThinkFilter() : undefined;
    let full = "";
    let started = false;
    const push = (raw: string) => {
      let t = filter ? filter.push(raw) : raw;
      if (!started) {
        t = t.replace(/^\s+/, "");
        if (!t) return;
        started = true;
      }
      if (!t) return;
      full += t;
      onText(t);
    };
    for await (const chunk of streamChat(ep, { model, messages: msgs, temperature, maxTokens: settings.maxTokens, responseFormat, signal })) {
      if (chunk.text) push(chunk.text);
      if (chunk.usage) usage = chunk.usage;
    }
    if (filter) push(filter.flush());
    return full;
  };

  let output = "";
  if (task.output.kind === "text") {
    output = await run(messages, undefined, (text) => emit({ type: "delta", text }));
  } else {
    const schema = task.output.schema;
    const jsonSchema = z.toJSONSchema(schema) as Record<string, unknown>;
    const mode = ep.structuredOutput;
    const responseFormat =
      mode === "json_schema"
        ? { type: "json_schema", json_schema: { name: task.output.schemaName, schema: jsonSchema } }
        : mode === "json_object"
          ? { type: "json_object" }
          : undefined;
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
    if (!parsed.ok) throw new UpstreamError(`Modellen returnerede ikke gyldigt struktureret svar (${parsed.error}). Prøv en større model.`, "UPSTREAM_PROTOCOL");
    emit({ type: "result", data: parsed.data });
  }

  emit({ type: "done", usage });

  if (settings.auditLog) {
    const entry = {
      ts: new Date().toISOString(),
      task: task.id,
      endpoint: new URL(ep.baseUrl).host,
      model,
      durationMs: Date.now() - started,
      inputChars: messages.reduce((n, m) => n + m.content.length, 0),
      outputChars: output.length,
      usage,
    };
    await appendFile(store.file.replace(/config\.json$/, "audit.log"), JSON.stringify(entry) + "\n").catch(() => {});
  }
}

function tryParse(schema: z.ZodType, text: string): { ok: true; data: unknown } | { ok: false; error: string } {
  try {
    const r = schema.safeParse(extractJson(text));
    return r.success ? { ok: true, data: r.data } : { ok: false, error: z.prettifyError(r.error).slice(0, 500) };
  } catch (err) {
    return { ok: false, error: (err as Error).message };
  }
}
