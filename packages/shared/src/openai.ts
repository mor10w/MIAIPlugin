import type { ChatMessage, Endpoint } from "./types.js";

/**
 * Minimal klient til OpenAI Chat Completions API – det eneste API, MIAI taler (ADR-0003).
 * Kører både i browseren (direkte forbindelse fra add-in'et) og i Node (bridgen), så den
 * bruger kun fetch, ReadableStream.getReader og AbortController.
 */

export type UpstreamCode = "UPSTREAM_UNREACHABLE" | "UPSTREAM_HTTP" | "UPSTREAM_TIMEOUT" | "UPSTREAM_PROTOCOL" | "NO_MODEL" | "POLICY";

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
  messages: ChatMessage[];
  temperature?: number;
  maxTokens?: number;
  presencePenalty?: number;
  responseFormat?: Record<string, unknown>;
  signal?: AbortSignal;
}

export interface StreamChunk {
  text?: string;
  usage?: { promptTokens?: number; completionTokens?: number };
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

export async function listModels(ep: Endpoint, signal?: AbortSignal, opts: ClientOptions = {}): Promise<string[]> {
  const res = await doFetch(ep, "/models", { method: "GET" }, signal, opts);
  const json = (await res.json().catch(() => null)) as { data?: { id?: string }[] } | null;
  if (!json || !Array.isArray(json.data)) throw new UpstreamError("Uventet svar fra /models", "UPSTREAM_PROTOCOL");
  return json.data.map((m) => m.id).filter((id): id is string => typeof id === "string");
}

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

  const res = await doFetch(ep, "/chat/completions", { method: "POST", body: JSON.stringify(body) }, p.signal, opts);
  const ctype = res.headers.get("content-type") ?? "";

  // Nogle servere ignorerer stream=true og svarer med ét JSON-objekt.
  if (!ctype.includes("text/event-stream")) {
    const json = (await res.json().catch(() => null)) as CompletionJson | null;
    const text = json?.choices?.[0]?.message?.content;
    if (typeof text !== "string") throw new UpstreamError("Uventet svar fra /chat/completions", "UPSTREAM_PROTOCOL");
    yield { text, usage: mapUsage(json?.usage) };
    return;
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

interface CompletionJson {
  choices?: { message?: { content?: string }; delta?: { content?: string | null } }[];
  usage?: { prompt_tokens?: number; completion_tokens?: number };
  error?: { message?: string } | string;
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
  const text = json.choices?.[0]?.delta?.content ?? undefined;
  const usage = mapUsage(json.usage);
  if (!text && !usage) return null;
  return { text: text ?? undefined, usage };
}
