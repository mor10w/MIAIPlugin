import type { Endpoint } from "./types.js";

/**
 * Modellens kontekst og kapabiliteter fra værktøjets eget API (#73). OpenAI's /v1/models oplyser
 * hverken konteksten (Ollama) eller den indlæste kontekst (LM Studio), så de læses her:
 *  - Ollama: POST /api/show (Modelfile-parametre, trænet maks., capabilities) og GET /api/ps
 *    (konteksten, en indlæst model faktisk kører med).
 *  - LM Studio: GET /api/v1/models (max_context_length, den indlæste instans' context_length,
 *    vision og trained_for_tool_use).
 * Kaldene går til samme host og port som endpointen, så politikken (allowedHosts) er uændret.
 */

export type ModelServer = "ollama" | "lmstudio";

/** Hvor kontekstlængden kommer fra. */
export type ContextSource =
  /** Angivet for endpointen i indstillingerne. */
  | "endpoint"
  /** Oplyst i /v1/models (vLLM m.fl.). */
  | "models"
  /** Ollama: PARAMETER num_ctx i modellens Modelfile. */
  | "ollama-modelfile"
  /** Ollama: den indlæste model (/api/ps). */
  | "ollama-loaded"
  /** LM Studio: den indlæste instans. */
  | "lmstudio-loaded";

export interface ModelDetails {
  model: string;
  /** Værktøjet bag endpointen, når det kunne genkendes. */
  server?: ModelServer;
  /** Konteksten, modellen kører med (tokens). Ukendt = undefined. */
  contextTokens?: number;
  contextSource?: ContextSource;
  /** Det højeste, modellen kan (trænet kontekst). */
  maxContextTokens?: number;
  vision?: boolean;
  tools?: boolean;
  thinking?: boolean;
}

/**
 * Under denne grænse advares der for agent og chat: systemprompt, dokument og historik fylder
 * typisk mere, og Ollama afkorter prompten uden at sige det.
 */
export const MIN_AGENT_CONTEXT = 16_384;

/** Ollamas standardkontekst på maskiner med under 24 GiB VRAM (docs.ollama.com/context-length). */
export const OLLAMA_SMALL_DEFAULT = 4_096;

/**
 * Hvordan konteksten gøres større i værktøjet (dansk; motorens fejlbeskeder er på dansk).
 * Ollamas OpenAI-lag ignorerer num_ctx i forespørgslen, så det skal sættes i Ollama selv.
 */
export function contextFix(server: ModelServer | undefined, prefix = ""): string {
  if (server === "ollama")
    return `${prefix}Ollama: sæt kontekstlængden i Ollama-appens indstillinger eller med miljøvariablen OLLAMA_CONTEXT_LENGTH (fx 32768), og genstart Ollama.`;
  if (server === "lmstudio") return `${prefix}LM Studio: indlæs modellen igen med en større Context Length (fx 32768).`;
  return "";
}

const positive = (v: unknown): number | undefined => (typeof v === "number" && Number.isFinite(v) && v > 0 ? Math.round(v) : undefined);

/** Ollamas navne: "qwen3" og "qwen3:latest" er samme model. */
function sameOllamaModel(a: string, b: string): boolean {
  const norm = (s: string) => (s.includes(":") ? s : `${s}:latest`);
  return norm(a) === norm(b);
}

/** Svaret fra Ollamas POST /api/show. */
export function parseOllamaShow(model: string, json: unknown): ModelDetails | undefined {
  if (!json || typeof json !== "object") return undefined;
  const j = json as { parameters?: unknown; model_info?: unknown; capabilities?: unknown };
  if (!j.model_info && !Array.isArray(j.capabilities)) return undefined;
  const info = (j.model_info && typeof j.model_info === "object" ? j.model_info : {}) as Record<string, unknown>;
  const arch = typeof info["general.architecture"] === "string" ? info["general.architecture"] : "";
  const trained = positive(info[`${arch}.context_length`]) ?? positive(Object.entries(info).find(([k]) => k.endsWith(".context_length"))?.[1]);
  const numCtx = typeof j.parameters === "string" ? positive(Number(/^\s*num_ctx\s+(\d+)/m.exec(j.parameters)?.[1])) : undefined;
  const caps = Array.isArray(j.capabilities) ? j.capabilities.filter((c): c is string => typeof c === "string") : undefined;
  return {
    model,
    server: "ollama",
    ...(numCtx ? { contextTokens: numCtx, contextSource: "ollama-modelfile" as const } : {}),
    maxContextTokens: trained,
    ...(caps ? { vision: caps.includes("vision"), tools: caps.includes("tools"), thinking: caps.includes("thinking") } : {}),
  };
}

/** Konteksten for modellen i Ollamas GET /api/ps, hvis den er indlæst. */
export function parseOllamaPs(model: string, json: unknown): number | undefined {
  const models = (json as { models?: unknown } | null)?.models;
  if (!Array.isArray(models)) return undefined;
  const hit = models.find((m) => {
    const name = (m as { model?: unknown; name?: unknown })?.model ?? (m as { name?: unknown })?.name;
    return typeof name === "string" && sameOllamaModel(name, model);
  }) as { context_length?: unknown } | undefined;
  return positive(hit?.context_length);
}

/** Svaret fra LM Studios GET /api/v1/models. */
export function parseLmStudioModels(model: string, json: unknown): ModelDetails | undefined {
  const models = (json as { models?: unknown } | null)?.models;
  if (!Array.isArray(models)) return undefined;
  type Lms = {
    key?: unknown;
    max_context_length?: unknown;
    loaded_instances?: { id?: unknown; config?: { context_length?: unknown } }[];
    capabilities?: { vision?: unknown; trained_for_tool_use?: unknown; reasoning?: unknown };
  };
  const list = models as Lms[];
  // OpenAI-laget bruger nøglen (fx "qwen/qwen3-8b") eller en indlæst instans' id (fx "qwen/qwen3-8b:2").
  const entry = list.find((m) => m?.key === model) ?? list.find((m) => Array.isArray(m?.loaded_instances) && m.loaded_instances.some((i) => i?.id === model));
  if (!entry) return { model, server: "lmstudio" };
  const instances = Array.isArray(entry.loaded_instances) ? entry.loaded_instances : [];
  const instance = instances.find((i) => i?.id === model) ?? instances[0];
  const loaded = positive(instance?.config?.context_length);
  const caps = entry.capabilities && typeof entry.capabilities === "object" ? entry.capabilities : undefined;
  return {
    model,
    server: "lmstudio",
    ...(loaded ? { contextTokens: loaded, contextSource: "lmstudio-loaded" as const } : {}),
    maxContextTokens: positive(entry.max_context_length),
    ...(caps
      ? {
          vision: caps.vision === true,
          tools: caps.trained_for_tool_use === true,
          ...(caps.reasoning ? { thinking: true } : {}),
        }
      : {}),
  };
}

/** Værktøjets rod-URL: endpointens adresse uden /v1. */
export function serverRoot(baseUrl: string): string {
  return baseUrl.replace(/\/+$/, "").replace(/\/v1$/i, "");
}

/** Hvilket værktøj prøves først: efter standardporten, ellers Ollama. */
function probeOrder(baseUrl: string): ModelServer[] {
  let port = "";
  try {
    port = new URL(baseUrl).port;
  } catch {
    /* ugyldig URL: prøv standardrækkefølgen */
  }
  return port === "1234" ? ["lmstudio", "ollama"] : ["ollama", "lmstudio"];
}

type Fetch = typeof fetch;

async function getJson(fetchFn: Fetch, url: string, init: RequestInit): Promise<unknown> {
  const res = await fetchFn(url, init);
  if (!res.ok) return undefined;
  return res.json().catch(() => undefined);
}

/** Servere, der hverken svarede som Ollama eller LM Studio: spørges ikke igen i 10 minutter. */
const notNative = new Map<string, number>();
const NOT_NATIVE_MS = 10 * 60_000;

/** Til tests og "Tjek igen": glem, hvilke servere der ikke var Ollama eller LM Studio. */
export function clearNativeCache(): void {
  notNative.clear();
}

/**
 * Læser modellens detaljer fra Ollamas eller LM Studios eget API. Returnerer undefined, hvis
 * værktøjet er et andet (eller ikke svarer); det er ikke en fejl, så kaldet må aldrig kaste.
 * `fresh` spørger igen, selv om serveren for nylig ikke svarede som et af dem.
 */
export async function fetchNativeModelInfo(
  ep: Pick<Endpoint, "baseUrl" | "apiKey">,
  model: string,
  opts: { signal?: AbortSignal; fetch?: Fetch; fresh?: boolean } = {},
): Promise<ModelDetails | undefined> {
  const fetchFn = opts.fetch ?? fetch;
  const root = serverRoot(ep.baseUrl);
  const miss = notNative.get(root);
  if (!opts.fresh && miss && Date.now() - miss < NOT_NATIVE_MS) return undefined;
  const auth: Record<string, string> = ep.apiKey ? { authorization: `Bearer ${ep.apiKey}` } : {};
  for (const server of probeOrder(ep.baseUrl)) {
    try {
      if (server === "ollama") {
        const show = parseOllamaShow(
          model,
          await getJson(fetchFn, `${root}/api/show`, {
            method: "POST",
            headers: { ...auth, "content-type": "application/json" },
            body: JSON.stringify({ model }),
            signal: opts.signal,
          }),
        );
        if (!show) continue;
        // En indlæst model kører med den kontekst, den blev indlæst med – også når Modelfile siger andet.
        const loaded = parseOllamaPs(model, await getJson(fetchFn, `${root}/api/ps`, { headers: auth, signal: opts.signal }).catch(() => undefined));
        return loaded ? { ...show, contextTokens: loaded, contextSource: "ollama-loaded" } : show;
      }
      const lms = parseLmStudioModels(model, await getJson(fetchFn, `${root}/api/v1/models`, { headers: auth, signal: opts.signal }));
      if (lms) return lms;
    } catch {
      if (opts.signal?.aborted) return undefined;
      // Prøv det næste værktøj.
    }
  }
  if (!opts.signal?.aborted) notNative.set(root, Date.now());
  return undefined;
}
