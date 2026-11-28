import { CELL_ASK_TASK, chosenModel, instructionsBlock, type Settings, type TaskEvent } from "@miai/shared";
import type { Backend } from "../backend";
import { makeT, type T } from "../i18n";
import { CellCache, queryKey } from "./cache";
import { CellCanceled, CellQueue, CellStopped, type QueueState } from "./queue";
import { cleanAnswer, inputToText, MAX_INPUT_CHARS, toCellValue } from "./value";

/**
 * Regnearksfunktionerne =MIAI.ASK() og =MIAI.MODELS() (#80, ADR-0018). De kører i samme side som
 * task pane (shared runtime) og går gennem den samme backend som opgaverne, så indstillinger,
 * politik (allowedHosts) og audit-log gælder uændret.
 */

/** En fejl, der vises i cellen: #N/A (kan ikke lige nu) eller #VALUE! (forkert brug). */
export class CellError extends Error {
  constructor(
    readonly kind: "na" | "value",
    message: string,
  ) {
    super(message);
    this.name = "CellError";
  }
}

/** Det, Excel giver funktionen som sidste argument. */
export interface CellInvocation {
  /** Sættes af funktionen; Excel kalder den, når kaldet afbrydes. */
  onCanceled?: () => void;
  /** Excel beregner kun en forhåndsvisning, mens formlen skrives (CustomFunctionsRuntime 1.5). */
  isInValuePreview?: boolean;
}

/** Vises i forhåndsvisningen, mens formlen skrives, når svaret ikke allerede er i cachen. */
export const PREVIEW_PLACEHOLDER = "…";

/**
 * Så længe venter et kald på, at MIAI selv svarer (backend, indstillinger, modelliste), før cellen
 * får en fejl. Selve modelkaldet har endpointens egen timeout. Uden grænsen står cellen på #BUSY!
 * uden forklaring, hvis noget i opstarten hænger.
 */
export const READY_TIMEOUT_MS = 20_000;

/** Afviser med `error`, hvis `p` ikke er afgjort inden for `ms`. */
function within<V>(p: V | Promise<V>, ms: number, error: () => Error): Promise<V> {
  return new Promise<V>((resolve, reject) => {
    const timer = setTimeout(() => reject(error()), ms);
    Promise.resolve(p).then(
      (v) => {
        clearTimeout(timer);
        resolve(v);
      },
      (e) => {
        clearTimeout(timer);
        reject(e);
      },
    );
  });
}

const MODELS_KEY = "miai.cellmodels.v1";
const SETTINGS_TTL_MS = 5_000;
const MODELS_TTL_MS = 30_000;

export interface CellRuntime {
  ask(prompt: unknown, input: unknown, optimized: unknown, model: unknown, invocation?: CellInvocation): Promise<string | number>;
  models(): Promise<string[][]>;
  /** Stopper alle kald i gang og i kø. */
  stop(): number;
  readonly state: QueueState;
  subscribe(fn: (s: QueueState) => void): () => void;
  readonly cacheSize: number;
  clearCache(): void;
}

export function createCellRuntime(backendReady: Backend | Promise<Backend>, cache = new CellCache(), readyTimeoutMs = READY_TIMEOUT_MS): CellRuntime {
  let settingsAt = 0;
  let settings: { settings: Settings; baseUrl: string; instructions: string } | undefined;
  let modelsAt = 0;
  let modelList: string[] | undefined;
  let concurrency = 2;
  const queue = new CellQueue(() => concurrency);
  // Skriv det, der venter, når siden lukkes (cachen samler ellers skrivningerne).
  if (typeof addEventListener === "function") addEventListener("pagehide", () => cache.flush());

  const loadSettings = async (backend: Backend) => {
    if (settings && Date.now() - settingsAt < SETTINGS_TTL_MS) return settings;
    const { settings: s, policy } = await backend.settings();
    const ep = s.endpoints.find((e) => e.id === s.activeEndpoint) ?? s.endpoints[0];
    // De faste instruktioner for Excel (#78) afgør også svaret: ændres de, er det en ny forespørgsel.
    settings = { settings: s, baseUrl: ep?.baseUrl ?? "", instructions: instructionsBlock("excel", policy?.instructions, s.instructions) ?? "" };
    settingsAt = Date.now();
    concurrency = s.cellConcurrency;
    return settings;
  };

  const loadModels = async (backend: Backend): Promise<string[]> => {
    if (modelList && Date.now() - modelsAt < MODELS_TTL_MS) return modelList;
    modelList = (await backend.models()).map((m) => m.id);
    modelsAt = Date.now();
    try {
      localStorage.setItem(MODELS_KEY, JSON.stringify(modelList));
    } catch {
      /* uden lager: listen lever kun i hukommelsen */
    }
    return modelList;
  };

  /** Fejl fra motoren → den fejl, cellen skal vise. */
  const fromEvent = (t: T, e: Extract<TaskEvent, { type: "error" }>): CellError => {
    switch (e.code) {
      case "POLICY":
        return new CellError("na", t("cellPolicy"));
      case "UPSTREAM_UNREACHABLE":
      case "UPSTREAM_TIMEOUT":
        return new CellError("na", t("cellUnreachable"));
      case "NO_MODEL":
        return new CellError("na", t("err_noModel"));
      case "CONTEXT_OVERFLOW":
        return new CellError("value", t("cellTooLong"));
      default:
        return new CellError("value", e.message.slice(0, 250));
    }
  };

  const callModel = async (
    backend: Backend,
    t: T,
    q: { prompt: string; input: string; optimized: boolean; model?: string },
    signal: AbortSignal,
  ): Promise<string> => {
    let text = "";
    let failed: CellError | undefined;
    await backend.runTask(
      CELL_ASK_TASK,
      {
        context: { host: "excel", text: q.input },
        // useKnowledge: bridgen slår kun op i egne mapper, hvis det også er slået til under Indstillinger.
        params: { prompt: q.prompt, optimized: q.optimized ? "on" : "", useKnowledge: q.optimized ? "" : "true" },
        model: q.model,
      },
      (e) => {
        if (e.type === "delta") text += e.text;
        // Tekst fra en runde, der endte med en web-søgning, trækkes tilbage.
        else if (e.type === "retract") text = text.slice(0, Math.max(0, text.length - e.chars));
        else if (e.type === "error") failed = fromEvent(t, e);
      },
      signal,
    );
    if (failed) throw failed;
    return cleanAnswer(text);
  };

  return {
    async ask(promptArg, inputArg, optimizedArg, modelArg, invocation) {
      let t = makeT("da");
      try {
        const notReady = () => new CellError("na", t("cellNotReady"));
        const backend = await within(backendReady, readyTimeoutMs, notReady);
        const { settings: s, baseUrl, instructions } = await within(loadSettings(backend), readyTimeoutMs, notReady);
        t = makeT(s.uiLanguage);

        const prompt = inputToText(promptArg).trim();
        if (!prompt) throw new CellError("value", t("cellNoPrompt"));
        const input = inputToText(inputArg);
        if (input.length > MAX_INPUT_CHARS) throw new CellError("value", t("cellTooLong"));
        const optimized = optimizedArg === true || optimizedArg === 1 || String(optimizedArg ?? "").toLowerCase() === "true";
        const model = inputToText(modelArg).trim() || undefined;

        const key = await queryKey({ endpoint: baseUrl, model: model ?? chosenModel(s, CELL_ASK_TASK) ?? "", optimized, prompt, input, instructions });
        const cached = cache.get(key);
        if (cached !== undefined) return toCellValue(cached);
        // Mens formlen skrives, kaldes modellen ikke.
        if (invocation?.isInValuePreview) return PREVIEW_PLACEHOLDER;

        if (model) {
          let known: string[];
          try {
            known = await within(loadModels(backend), readyTimeoutMs, () => new Error("timeout"));
          } catch {
            throw new CellError("na", t("cellUnreachable"));
          }
          if (!known.includes(model)) throw new CellError("value", t("cellUnknownModel"));
        }

        const ticket = queue.enqueue(key, (signal) => callModel(backend, t, { prompt, input, optimized, model }, signal));
        if (invocation) invocation.onCanceled = ticket.cancel;
        const answer = await ticket.promise;
        cache.set(key, answer);
        return toCellValue(answer);
      } catch (err) {
        if (err instanceof CellError) throw err;
        if (err instanceof CellStopped) throw new CellError("na", t("cellStopped"));
        // Afbrudt af Excel: resultatet bruges ikke, men løftet skal afsluttes.
        if (err instanceof CellCanceled) throw new CellError("na", t("cellStopped"));
        const code = (err as { code?: string })?.code;
        if (code === "TOKEN") throw new CellError("na", t("cellUnreachable"));
        throw new CellError("na", (err as Error)?.message?.slice(0, 250) || t("cellUnreachable"));
      }
    },

    async models() {
      const t = makeT("da");
      try {
        const list = await within(
          Promise.resolve(backendReady).then((b) => loadModels(b)),
          readyTimeoutMs,
          () => new Error("timeout"),
        );
        if (list.length) return list.map((m) => [m]);
      } catch {
        // Endpointen kan ikke nås: vis den sidst kendte liste, så en dropdown, der peger på den, ikke går i stykker.
        try {
          const saved = JSON.parse(localStorage.getItem(MODELS_KEY) ?? "[]") as unknown;
          if (Array.isArray(saved) && saved.length) return saved.filter((m): m is string => typeof m === "string").map((m) => [m]);
        } catch {
          /* intet gemt */
        }
        throw new CellError("na", t("cellUnreachable"));
      }
      throw new CellError("na", t("err_noModel"));
    },

    stop: () => queue.stopAll(),
    get state() {
      return queue.state;
    },
    subscribe: (fn) => queue.subscribe(fn),
    get cacheSize() {
      return cache.size;
    },
    clearCache: () => cache.clear(),
  };
}

interface CustomFunctionsApi {
  associate(id: string, fn: (...args: never[]) => unknown): void;
  Error?: new (code: string, message?: string) => Error;
  ErrorCode?: { notAvailable: string; invalidValue: string };
}

/** CellError → Excels egen fejltype (med forklaring i fejlmenuen), hvor den findes. */
function toExcelError(cf: CustomFunctionsApi, err: unknown): unknown {
  if (!(err instanceof CellError) || !cf.Error || !cf.ErrorCode) return err;
  return new cf.Error(err.kind === "na" ? cf.ErrorCode.notAvailable : cf.ErrorCode.invalidValue, err.message);
}

let registered: CellRuntime | undefined;

/**
 * Har manifestet regnearksfunktionerne med? Det kan siden ikke se på `CustomFunctions`: Excels
 * office.js definerer objektet i enhver task pane, også når manifestet ikke nævner funktionerne.
 * Manifester med funktionerne åbner derfor task pane med `?functions=1` (scripts/make-manifest.mjs).
 */
export function functionsDeclared(search: string = globalThis.location?.search ?? ""): boolean {
  return new URLSearchParams(search).has("functions");
}

/**
 * Knytter funktionerne til deres id'er i functions.json. Gør intet, når manifestet ikke har
 * funktionerne med (`declared`) eller uden for Excel, og kun én gang. Kaldes, før backend'en er
 * klar: Excel kan beregne celler, så snart siden er indlæst, og kaldene venter så på `backend`.
 */
export function registerCellFunctions(backend: Backend | Promise<Backend>, declared = functionsDeclared()): CellRuntime | undefined {
  if (registered) return registered;
  if (!declared) return undefined;
  const cf = (globalThis as { CustomFunctions?: CustomFunctionsApi }).CustomFunctions;
  if (!cf?.associate) return undefined;
  const rt = createCellRuntime(backend);
  const guard =
    <A extends unknown[], R>(fn: (...a: A) => Promise<R>) =>
    (...a: A) =>
      fn(...a).catch((e) => {
        throw toExcelError(cf, e);
      });
  cf.associate(
    "ASK",
    guard((prompt: unknown, input: unknown, optimized: unknown, model: unknown, invocation?: CellInvocation) =>
      rt.ask(prompt, input, optimized, model, invocation),
    ) as never,
  );
  cf.associate("MODELS", guard(() => rt.models()) as never);
  registered = rt;
  return rt;
}

/**
 * Som registerCellFunctions, men venter på, at Office.js har defineret `CustomFunctions`: det sker
 * først, når Office.js har hentet sin Excel-del, og ikke nødvendigvis før sidens egen kode kører.
 * Prøver i op til 30 sekunder og gør intet, når manifestet ikke har funktionerne med.
 */
export function registerCellFunctionsWhenReady(backend: Backend | Promise<Backend>, declared = functionsDeclared(), tries = 300): void {
  if (!declared || registerCellFunctions(backend, declared) || tries <= 0) return;
  setTimeout(() => registerCellFunctionsWhenReady(backend, declared, tries - 1), 100);
}

/** De registrerede funktioner, eller undefined uden for Excels shared runtime. */
export function cellRuntime(): CellRuntime | undefined {
  return registered;
}
