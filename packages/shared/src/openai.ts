import type { Endpoint } from "./types.js";

/**
 * Minimal klient til OpenAI Chat Completions API – det eneste API, MIAI taler (ADR-0003).
 * Kører både i browseren (direkte forbindelse fra add-in'et) og i Node (bridgen), så den
 * bruger kun fetch, ReadableStream.getReader og AbortController.
 */

export type UpstreamCode =
  "UPSTREAM_UNREACHABLE" | "UPSTREAM_HTTP" | "UPSTREAM_TIMEOUT" | "UPSTREAM_PROTOCOL" | "NO_MODEL" | "POLICY" | "CONTEXT_OVERFLOW" | "NO_EMBEDDINGS";

export class UpstreamError extends Error {
  constructor(
    message: string,
    readonly code: UpstreamCode,
    readonly status?: number,
  ) {
    super(message);
  }
}

export interface ClientOptions {
  /** Ekstra tekst ved forbindelsesfejl, fx en CORS-forklaring i direkte tilstand. */
  unreachableHint?: string;
}

export interface ChatParams {
  model: string;
  messages: ApiMessage[];
  temperature?: number;
  maxTokens?: number;
  presencePenalty?: number;
  responseFormat?: Record<string, unknown>;
  /** OpenAI-værktøjsdefinitioner (fx web_search). Se ADR-0010. */
  tools?: ToolSpec[];
  signal?: AbortSignal;
}

/** OpenAI-værktøjsdefinition (function calling). */
export interface ToolSpec {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, unknown> };
}

/** Et komplet værktøjskald fra modellen. `type` kræves af OpenAI-specifikationen, når kaldet sendes tilbage. */
export interface ToolCall {
  id: string;
  type?: "function";
  function: { name: string; arguments: string };
}

/** Delvist værktøjskald fra en stream-chunk (felterne bygges op over chunks). */
export interface ToolCallDelta {
  index: number;
  id?: string;
  function?: { name?: string; arguments?: string };
}

/** Besked til /chat/completions. Bredere end ChatMessage: tillader tool-kald og -resultater (ADR-0010). */
export type ApiMessage =
  { role: "system" | "user" | "assistant"; content: string; tool_calls?: ToolCall[] } | { role: "tool"; tool_call_id: string; content: string };

export interface StreamChunk {
  text?: string;
  usage?: { promptTokens?: number; completionTokens?: number };
  toolCalls?: ToolCallDelta[];
  finishReason?: string | null;
}

/** AbortSignal.any findes ikke i ældre WebViews (Safari < 17.4). */
function anySignal(signals: AbortSignal[]): AbortSignal {
  const anyFn = (AbortSignal as unknown as { any?: (s: AbortSignal[]) => AbortSignal }).any;
  if (anyFn) return anyFn(signals);
  const ctrl = new AbortController();
  for (const s of signals) {
    if (s.aborted) {
      ctrl.abort(s.reason);
      break;
    }
    s.addEventListener("abort", () => ctrl.abort(s.reason), { once: true });
  }
  return ctrl.signal;
}

function timeoutSignal(ms: number): AbortSignal {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(new Error("timeout")), ms) as unknown as { unref?: () => void };
  t.unref?.(); // Node: hold ikke processen i live
  return ctrl.signal;
}

function headers(ep: Endpoint, json: boolean): Record<string, string> {
  // Kun "simple" headers ved GET, så direkte kald fra browseren undgår unødige CORS-preflights.
  const h: Record<string, string> = {};
  if (json) h["content-type"] = "application/json";
  if (ep.apiKey) h.authorization = `Bearer ${ep.apiKey}`;
  return h;
}

function joinUrl(ep: Endpoint, path: string): string {
  return ep.baseUrl.replace(/\/+$/, "") + path;
}

async function doFetch(ep: Endpoint, path: string, init: RequestInit, signal: AbortSignal | undefined, opts: ClientOptions): Promise<Response> {
  const timeout = timeoutSignal(ep.timeoutMs);
  const combined = signal ? anySignal([signal, timeout]) : timeout;
  let res: Response;
  try {
    res = await fetch(joinUrl(ep, path), { ...init, headers: headers(ep, init.method === "POST"), signal: combined });
  } catch (err) {
    if (signal?.aborted) throw err;
    if (timeout.aborted) throw new UpstreamError(`Timeout efter ${ep.timeoutMs / 1000} s mod ${ep.baseUrl}`, "UPSTREAM_TIMEOUT");
    const hint = opts.unreachableHint ? ` ${opts.unreachableHint}` : "";
    throw new UpstreamError(`Kunne ikke forbinde til ${ep.baseUrl}. Kører dit inference-værktøj?${hint}`, "UPSTREAM_UNREACHABLE");
  }
  if (!res.ok) {
    const body = (await res.text().catch(() => "")).slice(0, 500);
    throw new UpstreamError(`${ep.name} svarede ${res.status}: ${body || res.statusText}`, "UPSTREAM_HTTP", res.status);
  }
  return res;
}

export interface ModelEntry {
  id: string;
  /** Modellens kontekstlængde, hvis værktøjet oplyser den (vLLM: max_model_len; andre: context_length m.fl.). */
  contextTokens?: number;
}

/** Felter, som forskellige værktøjer bruger til kontekstlængden i /v1/models. */
const CONTEXT_FIELDS = ["max_model_len", "context_length", "max_context_length", "n_ctx", "loaded_context_length"];

export async function listModelEntries(ep: Endpoint, signal?: AbortSignal, opts: ClientOptions = {}): Promise<ModelEntry[]> {
  const res = await doFetch(ep, "/models", { method: "GET" }, signal, opts);
  const json = (await res.json().catch(() => null)) as { data?: Record<string, unknown>[] } | null;
  if (!json || !Array.isArray(json.data)) throw new UpstreamError("Uventet svar fra /models", "UPSTREAM_PROTOCOL");
  return json.data
    .filter((m): m is Record<string, unknown> & { id: string } => typeof m.id === "string")
    .map((m) => {
      const ctx = CONTEXT_FIELDS.map((f) => m[f]).find((v) => typeof v === "number" && v > 0) as number | undefined;
      return { id: m.id, contextTokens: ctx };
    });
}

export async function listModels(ep: Endpoint, signal?: AbortSignal, opts: ClientOptions = {}): Promise<string[]> {
  return (await listModelEntries(ep, signal, opts)).map((m) => m.id);
}

/** Endpoints, der har afvist stream_options (så det ikke sendes igen i denne session). */
const noStreamOptions = new Set<string>();

/** Streamer en chat completion og giver tekst-deltaer. `reasoning_content` ignoreres bevidst. */
export async function* streamChat(ep: Endpoint, p: ChatParams, opts: ClientOptions = {}): AsyncGenerator<StreamChunk> {
  const body: Record<string, unknown> = {
    ...ep.extraBody,
    model: p.model,
    messages: p.messages,
    stream: true,
  };
  if (p.temperature !== undefined) body.temperature = p.temperature;
  if (p.maxTokens !== undefined) body.max_tokens = p.maxTokens;
  if (p.presencePenalty !== undefined) body.presence_penalty = p.presencePenalty;
  if (p.responseFormat) body.response_format = p.responseFormat;
  if (p.tools?.length) body.tools = p.tools;
  // Bed om usage i strømmen: bruges til at opdage, at modellen har afkortet prompten (#3).
  if (!noStreamOptions.has(ep.baseUrl)) body.stream_options = { include_usage: true };

  let res: Response;
  try {
    res = await doFetch(ep, "/chat/completions", { method: "POST", body: JSON.stringify(body) }, p.signal, opts);
  } catch (err) {
    // Ældre/strenge servere afviser ukendte felter. Prøv igen uden stream_options, og husk det.
    if (err instanceof UpstreamError && err.status === 400 && body.stream_options && /stream_options/i.test(err.message)) {
      noStreamOptions.add(ep.baseUrl);
      delete body.stream_options;
      res = await doFetch(ep, "/chat/completions", { method: "POST", body: JSON.stringify(body) }, p.signal, opts);
    } else throw err;
  }
  const ctype = res.headers.get("content-type") ?? "";

  // Nogle servere ignorerer stream=true og svarer med ét JSON-objekt.
  if (!ctype.includes("text/event-stream")) {
    const json = (await res.json().catch(() => null)) as CompletionJson | null;
    const choice = json?.choices?.[0];
    const msg = (choice as { message?: { content?: string; tool_calls?: RawToolCall[] } } | undefined)?.message;
    if (msg && typeof msg.content === "string") {
      yield { text: msg.content, usage: mapUsage(json?.usage), toolCalls: fullToolCalls(msg.tool_calls), finishReason: choice?.finish_reason ?? null };
      return;
    }
    if (msg?.tool_calls) {
      yield { toolCalls: fullToolCalls(msg.tool_calls), finishReason: choice?.finish_reason ?? "tool_calls", usage: mapUsage(json?.usage) };
      return;
    }
    throw new UpstreamError("Uventet svar fra /chat/completions", "UPSTREAM_PROTOCOL");
  }

  if (!res.body) throw new UpstreamError("Tomt svar fra /chat/completions", "UPSTREAM_PROTOCOL");
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = "";
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      let nl: number;
      while ((nl = buf.indexOf("\n")) >= 0) {
        const line = buf.slice(0, nl).replace(/\r$/, "");
        buf = buf.slice(nl + 1);
        const chunk = parseSseLine(line);
        if (chunk === "done") return;
        if (chunk) yield chunk;
      }
    }
    const last = parseSseLine(buf.trim());
    if (last && last !== "done") yield last;
  } finally {
    reader.cancel().catch(() => {});
  }
}

interface CompletionChoice {
  finish_reason?: string | null;
}

type Choice = CompletionChoice &
  ({ message?: { content?: string; tool_calls?: RawToolCall[] } } | { delta?: { content?: string | null; tool_calls?: RawToolCallDelta[] } });

interface CompletionJson {
  choices?: Choice[];
  usage?: { prompt_tokens?: number; completion_tokens?: number };
  error?: { message?: string } | string;
}

interface RawToolCall {
  id?: string;
  type?: string;
  function?: { name?: string; arguments?: string };
}

interface RawToolCallDelta extends RawToolCall {
  index?: number;
}

/** Samler værktøjskald ud af et svar (stream-delta eller komplet message). */
function fullToolCalls(raw: RawToolCall[] | undefined): ToolCallDelta[] | undefined {
  if (!raw?.length) return undefined;
  return raw.map((tc, i) => ({
    index: (tc as RawToolCallDelta).index ?? i,
    id: tc.id,
    function: { name: tc.function?.name, arguments: tc.function?.arguments ?? "" },
  }));
}

function mapUsage(u: CompletionJson["usage"]): StreamChunk["usage"] {
  return u ? { promptTokens: u.prompt_tokens, completionTokens: u.completion_tokens } : undefined;
}

export function parseSseLine(line: string): StreamChunk | "done" | null {
  if (!line.startsWith("data:")) return null;
  const data = line.slice(5).trim();
  if (!data) return null;
  if (data === "[DONE]") return "done";
  let json: CompletionJson;
  try {
    json = JSON.parse(data);
  } catch {
    return null;
  }
  if (json.error) {
    const msg = typeof json.error === "string" ? json.error : json.error.message;
    throw new UpstreamError(msg ?? "Ukendt fejl fra endpoint", "UPSTREAM_HTTP");
  }
  const choice = json.choices?.[0];
  const text = (choice as { delta?: { content?: string | null } } | undefined)?.delta?.content ?? undefined;
  const toolCalls = fullToolCalls((choice as { delta?: { tool_calls?: RawToolCallDelta[] } } | undefined)?.delta?.tool_calls);
  const usage = mapUsage(json.usage);
  if (!text && !usage && !toolCalls) return null;
  return { text: text ?? undefined, usage, toolCalls, finishReason: choice?.finish_reason ?? null };
}

// ---------------------------------------------------------------------------
// Embeddings (F15, lokal RAG – ADR-0013)
// ---------------------------------------------------------------------------

/**
 * Nogle embedding-modeller er trænet med faste præfikser for spørgsmål og dokumenter og giver
 * mærkbart dårligere søgning uden dem. Ukendte modeller får ingen præfiks.
 */
export function embeddingPrefix(model: string, kind: "query" | "document"): string {
  const m = model.toLowerCase();
  if (m.includes("nomic-embed")) return kind === "query" ? "search_query: " : "search_document: ";
  if (/(^|[/:_-])e5([-_:]|$)|multilingual-e5/.test(m)) return kind === "query" ? "query: " : "passage: ";
  if (m.includes("embeddinggemma")) return kind === "query" ? "task: search result | query: " : "title: none | text: ";
  if (m.includes("qwen3-embedding") && kind === "query")
    return "Instruct: Given a question, retrieve passages from the user's documents that answer it\nQuery: ";
  return "";
}

/**
 * Henter embeddings fra /v1/embeddings. Ikke alle OpenAI-kompatible værktøjer understøtter
 * endpointen (eller modellen er en chatmodel) – så kastes en fejl med koden NO_EMBEDDINGS og en
 * forklaring, brugeren kan handle på.
 */
export async function embed(ep: Endpoint, model: string, input: string[], signal?: AbortSignal): Promise<number[][]> {
  if (!model) throw new UpstreamError("Vælg en embedding-model under Indstillinger → Viden.", "NO_EMBEDDINGS");
  if (!input.length) return [];
  let res: Response;
  try {
    res = await doFetch(ep, "/embeddings", { method: "POST", body: JSON.stringify({ model, input }) }, signal, {});
  } catch (err) {
    const e = err as UpstreamError;
    if (e.code === "UPSTREAM_HTTP" && e.status !== undefined && [400, 404, 405, 501].includes(e.status))
      throw new UpstreamError(
        `${ep.name} kunne ikke lave embeddings med '${model}' (${e.message}). Tjek, at værktøjet understøtter /v1/embeddings, og at modellen er en embedding-model (fx bge-m3 eller nomic-embed-text).`,
        "NO_EMBEDDINGS",
        e.status,
      );
    throw err;
  }
  const json = (await res.json().catch(() => null)) as { data?: { embedding?: unknown; index?: number }[] } | null;
  const data = json?.data;
  if (!Array.isArray(data) || data.length !== input.length || !data.every((d) => Array.isArray(d.embedding) && (d.embedding as unknown[]).length > 0))
    throw new UpstreamError(`Uventet svar fra ${ep.name}/embeddings – er '${model}' en embedding-model?`, "NO_EMBEDDINGS");
  // Rækkefølgen følger `index`, hvis serveren angiver den.
  const sorted = data.every((d) => typeof d.index === "number") ? [...data].sort((a, b) => a.index! - b.index!) : data;
  return sorted.map((d) => d.embedding as number[]);
}
