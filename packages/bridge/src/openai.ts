import type { ChatMessage, Endpoint } from "@miai/shared";

/** Minimal klient til OpenAI Chat Completions API – det eneste API, bridgen taler (ADR-0003). */

export class UpstreamError extends Error {
  constructor(
    message: string,
    readonly code: "UPSTREAM_UNREACHABLE" | "UPSTREAM_HTTP" | "UPSTREAM_TIMEOUT" | "UPSTREAM_PROTOCOL" | "NO_MODEL",
    readonly status?: number,
  ) {
    super(message);
  }
}

export interface ChatParams {
  model: string;
  messages: ChatMessage[];
  temperature?: number;
  maxTokens?: number;
  responseFormat?: Record<string, unknown>;
  signal?: AbortSignal;
}

export interface StreamChunk {
  text?: string;
  usage?: { promptTokens?: number; completionTokens?: number };
}

function headers(ep: Endpoint): Record<string, string> {
  const h: Record<string, string> = { "content-type": "application/json", accept: "application/json, text/event-stream" };
  if (ep.apiKey) h.authorization = `Bearer ${ep.apiKey}`;
  return h;
}

function url(ep: Endpoint, path: string): string {
  return ep.baseUrl.replace(/\/+$/, "") + path;
}

async function doFetch(ep: Endpoint, path: string, init: RequestInit, signal?: AbortSignal): Promise<Response> {
  const timeout = AbortSignal.timeout(ep.timeoutMs);
  const combined = signal ? AbortSignal.any([signal, timeout]) : timeout;
  let res: Response;
  try {
    res = await fetch(url(ep, path), { ...init, headers: headers(ep), signal: combined });
  } catch (err) {
    if (timeout.aborted) throw new UpstreamError(`Timeout efter ${ep.timeoutMs / 1000} s mod ${ep.baseUrl}`, "UPSTREAM_TIMEOUT");
    if (signal?.aborted) throw err;
    throw new UpstreamError(`Kunne ikke forbinde til ${ep.baseUrl}. Kører dit inference-værktøj?`, "UPSTREAM_UNREACHABLE");
  }
  if (!res.ok) {
    const body = (await res.text().catch(() => "")).slice(0, 500);
    throw new UpstreamError(`${ep.name} svarede ${res.status}: ${body || res.statusText}`, "UPSTREAM_HTTP", res.status);
  }
  return res;
}

export async function listModels(ep: Endpoint, signal?: AbortSignal): Promise<string[]> {
  const res = await doFetch(ep, "/models", { method: "GET" }, signal);
  const json = (await res.json().catch(() => null)) as { data?: { id?: string }[] } | null;
  if (!json || !Array.isArray(json.data)) throw new UpstreamError("Uventet svar fra /models", "UPSTREAM_PROTOCOL");
  return json.data.map((m) => m.id).filter((id): id is string => typeof id === "string");
}

/** Streamer en chat completion og giver tekst-deltaer. `reasoning_content` ignoreres bevidst. */
export async function* streamChat(ep: Endpoint, p: ChatParams): AsyncGenerator<StreamChunk> {
  const body: Record<string, unknown> = {
    ...ep.extraBody,
    model: p.model,
    messages: p.messages,
    stream: true,
  };
  if (p.temperature !== undefined) body.temperature = p.temperature;
  if (p.maxTokens !== undefined) body.max_tokens = p.maxTokens;
  if (p.responseFormat) body.response_format = p.responseFormat;

  const res = await doFetch(ep, "/chat/completions", { method: "POST", body: JSON.stringify(body) }, p.signal);
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
  const decoder = new TextDecoder();
  let buf = "";
  for await (const bytes of res.body as unknown as AsyncIterable<Uint8Array>) {
    buf += decoder.decode(bytes, { stream: true });
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
