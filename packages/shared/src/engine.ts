import { z } from "zod";
import { hasImages, withAttachments } from "./attachments.js";
import { isDegenerating } from "./degenerate.js";
import { extractJson } from "./json.js";
import {
  listModelEntries,
  listModels,
  streamChat,
  UpstreamError,
  type ModelEntry,
  type ApiMessage,
  type ClientOptions,
  type ToolCall,
  type ToolSpec,
} from "./openai.js";
import { clearNativeCache, contextFix, fetchNativeModelInfo, MIN_AGENT_CONTEXT, type ModelDetails, type ModelServer } from "./modelInfo.js";
import { isUrlAllowed } from "./policy.js";
import { chunkChars, estimateMessages, looksTruncated, outputBudget, splitIntoChunks } from "./context.js";
import { doc, system, textFormat, type TaskDefinition } from "./tasks.js";
import { ThinkFilter } from "./think.js";
import { chosenModel } from "./modelGroups.js";
import {
  agentExportAllowed,
  agentScriptsAllowed,
  type Endpoint,
  type ParsedTaskRequest,
  type Policy,
  type ProbeResult,
  type Settings,
  type TaskEvent,
  type TaskWarning,
} from "./types.js";

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

/** Map-trinnet ved opdeling af lange tekster (#3). */
const MAP_INSTRUCTION =
  "This is one part of a longer document. Summarize this part concisely: keep key facts, numbers, names, decisions and conclusions. Do not add an introduction. Answer in the same language as the text.";
const REDUCE_NOTE =
  "The document was too long to read at once. The material below consists of summaries of its consecutive parts, in order. Treat it as the whole document.";

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

const modelCache = new Map<string, { entries: ModelEntry[]; at: number }>();

async function cachedEntries(ep: Endpoint, signal?: AbortSignal, opts?: ClientOptions): Promise<ModelEntry[]> {
  const key = `${ep.baseUrl}|${ep.apiKey}`;
  const hit = modelCache.get(key);
  if (hit && Date.now() - hit.at < 30_000) return hit.entries;
  const entries = await listModelEntries(ep, signal, opts);
  modelCache.set(key, { entries, at: Date.now() });
  return entries;
}

export async function cachedModels(ep: Endpoint, signal?: AbortSignal, opts?: ClientOptions): Promise<string[]> {
  return (await cachedEntries(ep, signal, opts)).map((m) => m.id);
}

const detailsCache = new Map<string, { details: ModelDetails; at: number }>();

/**
 * Modellens kontekst og kapabiliteter (#73). Konteksten tages i denne rækkefølge: endpointens
 * indstilling, /v1/models (vLLM m.fl.), værktøjets eget API (Ollama, LM Studio). Kapabiliteterne
 * kommer kun fra værktøjets eget API. Ukendt kontekst (fx Ollama med en model, der ikke er indlæst
 * endnu) → undefined; så planlægges der uden loft. `fresh` springer cachen over, fx efter en
 * kørsel, der har indlæst modellen.
 */
export async function modelDetails(
  ep: Endpoint,
  model: string,
  opts: { signal?: AbortSignal; client?: ClientOptions; fresh?: boolean } = {},
): Promise<ModelDetails> {
  const key = `${ep.baseUrl}|${ep.apiKey}|${model}`;
  const hit = detailsCache.get(key);
  if (!opts.fresh && hit && Date.now() - hit.at < 30_000) return withEndpointContext(ep, hit.details);
  const native = await fetchNativeModelInfo(ep, model, { signal: opts.signal, fresh: opts.fresh });
  let fromModels: number | undefined;
  if (!native?.contextTokens) {
    try {
      fromModels = (await cachedEntries(ep, opts.signal, opts.client)).find((m) => m.id === model)?.contextTokens;
    } catch {
      /* ukendt */
    }
  }
  const details: ModelDetails = fromModels ? { ...native, model, contextTokens: fromModels, contextSource: "models" } : (native ?? { model });
  detailsCache.set(key, { details, at: Date.now() });
  return withEndpointContext(ep, details);
}

/** Endpointens egen indstilling har forrang for det aflæste. */
function withEndpointContext(ep: Endpoint, d: ModelDetails): ModelDetails {
  return ep.contextTokens ? { ...d, contextTokens: ep.contextTokens, contextSource: "endpoint" } : d;
}

/** Modellens kontekstlængde, eller undefined, når den ikke kendes (se modelDetails). */
export async function contextTokensFor(ep: Endpoint, model: string, signal?: AbortSignal, opts?: ClientOptions): Promise<number | undefined> {
  if (ep.contextTokens) return ep.contextTokens;
  try {
    return (await modelDetails(ep, model, { signal, client: opts })).contextTokens;
  } catch {
    return undefined;
  }
}

export function clearModelCache(): void {
  modelCache.clear();
  detailsCache.clear();
  clearNativeCache();
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
/**
 * Billeder til en model, der ikke kan se (#27): de fleste værktøjer afviser forespørgslen med en
 * HTTP-fejl, som ikke siger brugeren noget. Den oversættes til en forklaring med en vej videre.
 */
async function* withVisionError<T>(stream: AsyncGenerator<T>, model: string | undefined): AsyncGenerator<T> {
  try {
    yield* stream;
  } catch (err) {
    if (model && err instanceof UpstreamError && err.code === "UPSTREAM_HTTP" && err.status !== 401 && err.status !== 403 && err.status !== 429)
      throw new UpstreamError(
        `${model} kunne ikke bruge det vedhæftede billede: ${err.message}. Modellen understøtter måske ikke billeder – vælg en model med vision, eller fjern billedet.`,
        "NO_VISION",
        err.status,
      );
    throw err;
  }
}

export async function executeTask(task: TaskDefinition, req: ParsedTaskRequest, deps: EngineDeps): Promise<TaskSummary> {
  const { settings, endpoint: ep, policy, emit, signal, client } = deps;
  if (!isUrlAllowed(ep.baseUrl, policy)) {
    throw new UpstreamError(`Endpoint ${ep.baseUrl} er ikke tilladt af politikken.`, "POLICY");
  }

  // Opgavens eget valg, så dens type, så standardmodellen, så den første fra endpointen.
  const model = req.model || chosenModel(settings, task.id) || (await cachedModels(ep, signal, client))[0];
  if (!model) throw new UpstreamError(`${ep.name} har ingen modeller. Hent eller indlæs en model i dit inference-værktøj.`, "NO_MODEL");
  emit({ type: "meta", model, endpoint: ep.name });

  const details = await modelDetails(ep, model, { signal, client });
  const contextTokens = details.contextTokens;
  const warnings = new Set<TaskWarning>();
  let inputReq = req;

  // Map-reduce (#3): lange tekster til opgaver, der understøtter det (fx opsummering), deles
  // ved afsnitsgrænser; hver del opsummeres, og den afsluttende kørsel arbejder på del-resuméerne.
  const mapReduce = task.chunking === "map-reduce" && req.context.text.length > chunkChars(contextTokens);

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
  // Opgavens egen justering (#80): fx regnearksfunktionen med optimized = SAND.
  const tuning = task.tuning?.(req) ?? {};
  const tools = deps.tools && settings.webSearch && policy.webSearch && !tuning.noTools ? deps.tools : undefined;
  const maxTokensSetting = tuning.maxTokens ? Math.min(settings.maxTokens, tuning.maxTokens) : settings.maxTokens;
  const thinking = tuning.noThinking ? "off" : settings.thinking === "default" ? undefined : settings.thinking;

  const addUsage = (u: { promptTokens?: number; completionTokens?: number } | undefined) => {
    if (!u) return;
    usage = {
      promptTokens: (usage?.promptTokens ?? 0) + (u.promptTokens ?? 0),
      completionTokens: (usage?.completionTokens ?? 0) + (u.completionTokens ?? 0),
    };
  };

  const run = async (
    msgs: ApiMessage[],
    responseFormat: Record<string, unknown> | undefined,
    onText: (t: string) => void,
    useTools = true,
  ): Promise<{ text: string; toolCalls?: ToolCall[] }> => {
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
    // Kontekstbudget (#3): med kendt kontekst afvises alt for lange prompts, og svaret begrænses,
    // så prompt + svar passer (vLLM m.fl. afviser ellers forespørgslen).
    const promptEstimate = estimateMessages(msgs);
    if (contextTokens && promptEstimate > contextTokens * 1.1) {
      throw new UpstreamError(
        `Teksten er for lang til modellens kontekst (ca. ${promptEstimate} tokens; modellen kan ${contextTokens}). Markér mindre tekst, eller vælg en model med større kontekst.${contextFix(details.server, " ")}`,
        "CONTEXT_OVERFLOW",
      );
    }
    const maxTokens = outputBudget(maxTokensSetting, promptEstimate, contextTokens);
    let reqUsage: { promptTokens?: number } | undefined;
    let finishReason: string | null | undefined;
    // Værktøjskald bygges op over stream-chunks (arguments ankommer som fragmenter).
    const tcParts = new Map<number, { id: string; name: string; args: string }>();
    for await (const chunk of withVisionError(
      streamChat(
        ep,
        {
          model,
          messages: msgs,
          temperature,
          maxTokens,
          presencePenalty,
          seed: task.seed,
          responseFormat,
          signal,
          tools: useTools ? tools?.specs : undefined,
          thinking,
          reasoningEffort: tuning.noThinking || settings.reasoningEffort === "default" ? undefined : settings.reasoningEffort,
        },
        client,
      ),
      hasImages(req.attachments) ? `${ep.name} (${model})` : undefined,
    )) {
      if (chunk.thinkingIgnored) warnings.add("thinking-ignored");
      if (chunk.text) push(chunk.text);
      addUsage(chunk.usage);
      if (chunk.usage) reqUsage = chunk.usage;
      if (chunk.finishReason) finishReason = chunk.finishReason;
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
    // Modellen så mindre, end der blev sendt (fx Ollama over num_ctx) – eller svaret ramte loftet.
    if (looksTruncated(promptEstimate, reqUsage?.promptTokens)) warnings.add("input-truncated");
    if (finishReason === "length" && stopReason !== "repetition") warnings.add("output-truncated");
    const toolCalls = tcParts.size
      ? [...tcParts.entries()].sort((a, b) => a[0] - b[0]).map(([i, p]) => ({ id: p.id || `call_${i}`, function: { name: p.name, arguments: p.args } }))
      : undefined;
    return { text: full, toolCalls };
  };

  if (mapReduce) {
    const parts = splitIntoChunks(req.context.text, chunkChars(contextTokens));
    const partials: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      emit({ type: "progress", chars: 0, step: { done: i, total: parts.length } });
      const r = await run([system(req, MAP_INSTRUCTION), { role: "user", content: doc(parts[i]) }], undefined, () => {}, false);
      partials.push(`Part ${i + 1} of ${parts.length}:\n${r.text}`);
    }
    emit({ type: "progress", chars: 0, step: { done: parts.length, total: parts.length } });
    inputReq = { ...req, context: { ...req.context, text: partials.join("\n\n") } };
  }
  // Kodetilstand (ADR-0017): agenten får kun script-værktøjet at vide, når brugeren har slået det
  // til, og politikken tillader det. Det afgøres her, så bridge og direkte tilstand er ens.
  // Eksport (tillæg ADR-0012): kun bridgen kan skrive filer – direkte tilstand har ingen eksekverer.
  const agentExport = deps.tools !== undefined && agentExportAllowed(settings, policy) ? "on" : "";
  inputReq = { ...inputReq, params: { ...inputReq.params, agentScripts: agentScriptsAllowed(settings, policy) ? "on" : "", agentExport } };
  const built = task.buildMessages(inputReq);
  // Vedhæftninger (#27) lægges på brugerens seneste besked: tekstfiler som data, billeder som billeddele.
  const messages: ApiMessage[] = withAttachments(
    mapReduce ? [{ ...built[0], content: `${built[0].content}\n${REDUCE_NOTE}` }, ...built.slice(1)] : built,
    req.attachments,
  );

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
      // Tekst fra en runde, der ender med en søgning (fx "Lad mig søge…"), hører ikke til svaret.
      if (r.text) emit({ type: "retract", chars: r.text.length });
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
    if (!parsed.ok)
      throw new UpstreamError(`Modellen returnerede ikke et gyldigt struktureret svar (${parsed.error}). Prøv en større model.`, "UPSTREAM_PROTOCOL");
    emit({ type: "result", data: parsed.data });
  }

  // Lille kontekst (#73): agent og chat har brug for plads til systemprompt, dokument og historik.
  // Ollama kender først konteksten, når modellen er indlæst – det er den nu, så spørg igen.
  let context: { tokens: number; server?: ModelServer } | undefined;
  if (task.id === "agent" || task.id === "chat") {
    const d = contextTokens || details.server !== "ollama" ? details : await modelDetails(ep, model, { signal, client, fresh: true });
    if (d.contextTokens && d.contextTokens < MIN_AGENT_CONTEXT) {
      warnings.add("small-context");
      context = { tokens: d.contextTokens, ...(d.server ? { server: d.server } : {}) };
    }
  }

  emit({ type: "done", usage, stopReason, warnings: warnings.size ? [...warnings] : undefined, ...(context ? { context } : {}) });
  return {
    model,
    output,
    usage,
    durationMs: Date.now() - started,
    // Tegn sendt til modellen: den afsluttende kørsel + ved map-reduce også alle dele.
    inputChars: messages.reduce((n, m) => n + (m.content?.length ?? 0), 0) + (mapReduce ? req.context.text.length : 0),
  };
}

/** Omsætter en fejl til en TaskEvent. */
export function errorEvent(err: unknown): TaskEvent {
  const code = err instanceof UpstreamError ? err.code : err instanceof InputError ? "NO_INPUT" : "INTERNAL";
  const status = err instanceof UpstreamError ? err.status : undefined;
  return { type: "error", message: (err as Error)?.message ?? String(err), code, ...(status ? { status } : {}) };
}

function tryParse(schema: z.ZodType, text: string): { ok: true; data: unknown } | { ok: false; error: string } {
  try {
    const r = schema.safeParse(extractJson(text));
    return r.success ? { ok: true, data: r.data } : { ok: false, error: z.prettifyError(r.error).slice(0, 500) };
  } catch (err) {
    return { ok: false, error: (err as Error).message };
  }
}

/**
 * Tester en endpoint uden at gemme den (#4): svarer /v1/models, og hvilke modeller findes?
 * `reachableWithoutCors` bruges i browseren til at skelne "CORS mangler" fra "intet kører":
 * et kald med mode "no-cors" lykkes (opakt svar), når noget lytter på porten.
 */
export async function probeEndpoint(
  ep: Endpoint,
  policy: Policy,
  opts: { signal?: AbortSignal; client?: ClientOptions; reachableWithoutCors?: (url: string) => Promise<boolean> } = {},
): Promise<ProbeResult> {
  if (!isUrlAllowed(ep.baseUrl, policy)) return { ok: false, reason: "policy", message: `${ep.baseUrl} er ikke tilladt af politikken.` };
  // Kort timeout: en lukket port svarer straks, men en filtreret kan hænge.
  const probeEp = { ...ep, timeoutMs: Math.min(ep.timeoutMs, 5_000) };
  try {
    return { ok: true, models: await listModels(probeEp, opts.signal, opts.client) };
  } catch (err) {
    if (!(err instanceof UpstreamError)) throw err;
    if (err.code === "UPSTREAM_UNREACHABLE" || err.code === "UPSTREAM_TIMEOUT") {
      const url = ep.baseUrl.replace(/\/+$/, "") + "/models";
      if (opts.reachableWithoutCors && (await opts.reachableWithoutCors(url).catch(() => false))) {
        return { ok: false, reason: "cors", message: err.message };
      }
      return { ok: false, reason: "absent", message: err.message };
    }
    return { ok: false, reason: err.code === "UPSTREAM_HTTP" ? "http" : "protocol", message: err.message };
  }
}
