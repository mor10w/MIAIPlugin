import { defaultSettings, Settings, type TaskEvent, type TaskRequest } from "@miai/shared";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Backend } from "../backend";
import { CellCache, type CacheStorage } from "./cache";
import { CellError, createCellRuntime, functionsDeclared, PREVIEW_PLACEHOLDER, registerCellFunctions, type CellInvocation } from "./functions";

const memory = (): CacheStorage => {
  const data = new Map<string, string>();
  return { get: (k) => data.get(k) ?? null, set: (k, v) => void data.set(k, v), remove: (k) => void data.delete(k) };
};

/** En backend, hvor testen bestemmer svaret pr. kald. */
function fakeBackend(over: Partial<Settings> = {}) {
  const calls: { id: string; req: TaskRequest; signal: AbortSignal }[] = [];
  const state = {
    settings: Settings.parse({ ...defaultSettings(), ...over }),
    models: ["qwen3:4b", "qwen3:32b"],
    modelsFail: false,
    reply: (req: TaskRequest): TaskEvent[] | Promise<TaskEvent[]> => [{ type: "delta", text: `svar på ${req.context.text}` }, { type: "done" }],
  };
  const backend = {
    mode: "direct",
    settings: async () => ({ settings: state.settings, policy: {}, policySource: "default" }),
    models: async () => {
      if (state.modelsFail) throw new Error("intet net");
      return state.models.map((id) => ({ id, endpoint: "ollama" }));
    },
    runTask: async (id: string, req: TaskRequest, onEvent: (e: TaskEvent) => void, signal: AbortSignal) => {
      calls.push({ id, req, signal });
      const events = await Promise.race([
        Promise.resolve(state.reply(req)),
        new Promise<never>((_, fail) => signal.addEventListener("abort", () => fail(new Error("aborted")))),
      ]);
      for (const e of events) onEvent(e);
    },
  } as unknown as Backend;
  return { backend, calls, state };
}

let f: ReturnType<typeof fakeBackend>;
let rt: ReturnType<typeof createCellRuntime>;
beforeEach(() => {
  f = fakeBackend();
  rt = createCellRuntime(f.backend, new CellCache(memory()));
});

describe("=MIAI.ASK()", () => {
  it("sender prompt og input til opgaven cell-ask og returnerer svaret", async () => {
    expect(await rt.ask("Oversæt til engelsk", [["Hej"]], null, null)).toBe("svar på Hej");
    expect(f.calls).toHaveLength(1);
    expect(f.calls[0].id).toBe("cell-ask");
    expect(f.calls[0].req.context).toMatchObject({ host: "excel", text: "Hej" });
    expect(f.calls[0].req.params).toEqual({ prompt: "Oversæt til engelsk", optimized: "", useKnowledge: "true" });
    expect(f.calls[0].req.model).toBeUndefined();
  });

  it("et område sendes som TSV, og prompten kan komme fra en celle", async () => {
    await rt.ask(
      [["Find byen"]],
      [
        ["Acme", "Aarhus"],
        ["Beta", "Odense"],
      ],
      null,
      null,
    );
    expect(f.calls[0].req.context.text).toBe("Acme\tAarhus\nBeta\tOdense");
    expect(f.calls[0].req.params?.prompt).toBe("Find byen");
  });

  it("optimized = SAND giver den optimerede profil og en anden cachenøgle", async () => {
    await rt.ask("p", [["x"]], true, null);
    expect(f.calls[0].req.params).toMatchObject({ optimized: "on", useKnowledge: "" });
    await rt.ask("p", [["x"]], false, null);
    expect(f.calls).toHaveLength(2);
  });

  it("uændrede celler bruger cachen; ændret input kalder modellen igen", async () => {
    await rt.ask("p", [["x"]], null, null);
    await rt.ask("p", [["x"]], null, null);
    expect(f.calls).toHaveLength(1);
    await rt.ask("p", [["y"]], null, null);
    expect(f.calls).toHaveLength(2);
    rt.clearCache();
    expect(rt.cacheSize).toBe(0);
    await rt.ask("p", [["x"]], null, null);
    expect(f.calls).toHaveLength(3);
  });

  it("tal returneres som tal, postnumre og formler som tekst, og kodehegn fjernes", async () => {
    const answer = async (text: string) => {
      f.state.reply = () => [{ type: "delta", text }, { type: "done" }];
      return rt.ask(`p ${text}`, null, null, null);
    };
    expect(await answer("1234,5")).toBe(1234.5);
    expect(await answer("0045")).toBe("0045");
    expect(await answer("2300 København")).toBe("2300 København");
    expect(await answer("=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
    expect(await answer("```\nDanmark\n```")).toBe("Danmark");
  });

  it("tekst fra en runde, der endte med en websøgning, trækkes tilbage", async () => {
    f.state.reply = () => [{ type: "delta", text: "Jeg søger…" }, { type: "retract", chars: 10 }, { type: "delta", text: "København" }, { type: "done" }];
    expect(await rt.ask("Hovedstad?", [["Danmark"]], null, null)).toBe("København");
  });

  it("model: bruges, når den findes – også når optimized er sprunget over; ukendt navn giver #VALUE! uden modelkald", async () => {
    await rt.ask("p", [["x"]], null, "qwen3:32b");
    expect(f.calls[0].req.model).toBe("qwen3:32b");
    const err = await rt.ask("p", [["x"]], null, "findes-ikke").catch((e) => e);
    expect(err).toBeInstanceOf(CellError);
    expect(err).toMatchObject({ kind: "value", message: "Modellen findes ikke på endpointen" });
    expect(f.calls).toHaveLength(1);
    // Modelnavnet indgår i cachenøglen.
    await rt.ask("p", [["x"]], null, "qwen3:4b");
    expect(f.calls).toHaveLength(2);
  });

  it("mens formlen skrives, kaldes modellen ikke – men et svar i cachen vises", async () => {
    const preview: CellInvocation = { isInValuePreview: true };
    expect(await rt.ask("p", [["x"]], null, null, preview)).toBe(PREVIEW_PLACEHOLDER);
    expect(f.calls).toHaveLength(0);
    await rt.ask("p", [["x"]], null, null);
    expect(await rt.ask("p", [["x"]], null, null, preview)).toBe("svar på x");
  });

  it("fejl bliver til #N/A eller #VALUE! med forklaring, og fejl caches ikke", async () => {
    expect(await rt.ask("  ", [["x"]], null, null).catch((e) => e)).toMatchObject({ kind: "value", message: "Prompten mangler" });
    const fail = async (code: string) => {
      f.state.reply = () => [{ type: "error", message: "rå fejl", code }];
      return rt.ask(`p ${code}`, [["x"]], null, null).catch((e) => e);
    };
    expect(await fail("UPSTREAM_UNREACHABLE")).toMatchObject({ kind: "na", message: "MIAI kan ikke nå modellen" });
    expect(await fail("POLICY")).toMatchObject({ kind: "na", message: "Ikke tilladt af politikken" });
    expect(await fail("CONTEXT_OVERFLOW")).toMatchObject({ kind: "value", message: "Input er for langt til modellens kontekst" });
    expect(await fail("UPSTREAM_PROTOCOL")).toMatchObject({ kind: "value", message: "rå fejl" });
    const before = f.calls.length;
    f.state.reply = () => [{ type: "delta", text: "nu virker det" }, { type: "done" }];
    expect(await rt.ask("p POLICY", [["x"]], null, null)).toBe("nu virker det");
    expect(f.calls).toHaveLength(before + 1);
  });

  it("fejlbeskederne følger sproget i indstillingerne", async () => {
    f = fakeBackend({ uiLanguage: "en" });
    rt = createCellRuntime(f.backend, new CellCache(memory()));
    expect(await rt.ask("", null, null, null).catch((e) => e)).toMatchObject({ message: "The prompt is missing" });
  });

  it("loftet over samtidige kald følger indstillingen, Stop giver 'Stoppet', og Excels afbrydelse afbryder kaldet", async () => {
    f = fakeBackend({ cellConcurrency: 2 });
    rt = createCellRuntime(f.backend, new CellCache(memory()));
    f.state.reply = () => new Promise(() => {}); // modellen svarer aldrig
    const states: number[] = [];
    rt.subscribe((s) => states.push(s.running));
    const pending = [1, 2, 3, 4, 5].map((i) => rt.ask("p", [[`række ${i}`]], true, null).catch((e) => e));
    await new Promise((r) => setTimeout(r, 20));
    expect(rt.state).toEqual({ running: 2, queued: 3 });
    expect(Math.max(...states)).toBe(2);
    expect(rt.stop()).toBe(5);
    for (const e of await Promise.all(pending)) expect(e).toMatchObject({ kind: "na", message: "Stoppet" });
    expect(f.calls.every((c) => c.signal.aborted)).toBe(true);

    const invocation: CellInvocation = {};
    const one = rt.ask("p", [["afbrydes"]], null, null, invocation).catch((e) => e);
    await new Promise((r) => setTimeout(r, 20));
    invocation.onCanceled!();
    expect(await one).toBeInstanceOf(CellError);
    expect(f.calls.at(-1)!.signal.aborted).toBe(true);
  });
});

describe("funktionen hænger aldrig uden forklaring", () => {
  it("bliver MIAI aldrig klar, får cellen #N/A med en forklaring i stedet for at stå på #BUSY!", async () => {
    const never = new Promise<Backend>(() => {});
    const stuck = createCellRuntime(never, new CellCache(memory()), 30);
    expect(await stuck.ask("p", [["x"]], true, null).catch((e) => e)).toMatchObject({
      kind: "na",
      message: "MIAI svarer ikke. Åbn MIAI-panelet, tjek forbindelsen, og beregn cellen igen.",
    });
    expect(await stuck.models().catch((e) => e)).toMatchObject({ kind: "na" });
  });
  it("svarer indstillingerne aldrig, gælder det samme", async () => {
    const slow = { ...f.backend, settings: () => new Promise(() => {}) } as unknown as Backend;
    const rt2 = createCellRuntime(slow, new CellCache(memory()), 30);
    expect(await rt2.ask("p", [["x"]], true, null).catch((e) => e)).toMatchObject({ kind: "na" });
    expect(f.calls).toHaveLength(0);
  });
  it("en backend, der først bliver klar lidt senere, er ikke en fejl", async () => {
    const later = new Promise<Backend>((ok) => setTimeout(() => ok(f.backend), 20));
    const rt2 = createCellRuntime(later, new CellCache(memory()), 2_000);
    expect(await rt2.ask("p", [["x"]], true, null)).toBe("svar på x");
  });
});

describe("=MIAI.MODELS()", () => {
  it("giver modellerne som en kolonne; en endpoint uden modeller giver #N/A", async () => {
    expect(await rt.models()).toEqual([["qwen3:4b"], ["qwen3:32b"]]);
    f.state.models = [];
    const fresh = createCellRuntime(f.backend, new CellCache(memory()));
    expect(await fresh.models().catch((e) => e)).toMatchObject({ kind: "na" });
  });
});

describe("registrering", () => {
  const fakeApi = (fns: Record<string, (...a: unknown[]) => Promise<unknown>> = {}) => ({
    associate: (id: string, fn: (...a: unknown[]) => Promise<unknown>) => (fns[id] = fn),
  });

  it("manifestet fortæller siden, om funktionerne er med (?functions=1)", () => {
    expect(functionsDeclared("?functions=1")).toBe(true);
    expect(functionsDeclared("?token=abc&functions=1")).toBe(true);
    expect(functionsDeclared("?token=abc")).toBe(false);
    expect(functionsDeclared("")).toBe(false);
  });
  it("gør intet uden for Excel, hvor CustomFunctions ikke findes", () => {
    expect(registerCellFunctions(f.backend, true)).toBeUndefined();
  });
  it("gør intet, når manifestet ikke har funktionerne med – selv om Excels office.js altid definerer CustomFunctions", () => {
    const fns = {};
    (globalThis as Record<string, unknown>).CustomFunctions = fakeApi(fns);
    try {
      expect(registerCellFunctions(f.backend, false)).toBeUndefined();
      expect(registerCellFunctions(f.backend)).toBeUndefined(); // ingen ?functions=1 i adressen
      expect(Object.keys(fns)).toEqual([]);
    } finally {
      delete (globalThis as Record<string, unknown>).CustomFunctions;
    }
  });
  it("venter på, at Office.js definerer CustomFunctions, når det sker efter sidens egen kode", async () => {
    const fns: Record<string, (...a: unknown[]) => Promise<unknown>> = {};
    // Registreringen sker kun én gang pr. side; en frisk kopi af modulet svarer til en ny side.
    vi.resetModules();
    const { registerCellFunctionsWhenReady } = await import("./functions");
    try {
      registerCellFunctionsWhenReady(f.backend, false); // manifestet har ikke funktionerne med: intet sker
      registerCellFunctionsWhenReady(f.backend, true);
      expect(Object.keys(fns)).toEqual([]);
      setTimeout(() => ((globalThis as Record<string, unknown>).CustomFunctions = fakeApi(fns)), 150);
      await new Promise((r) => setTimeout(r, 450));
      expect(Object.keys(fns).sort()).toEqual(["ASK", "MODELS"]);
    } finally {
      delete (globalThis as Record<string, unknown>).CustomFunctions;
    }
  });
  it("knytter ASK og MODELS til, og fejl bliver til Excels egen fejltype med besked", async () => {
    const fns: Record<string, (...a: unknown[]) => Promise<unknown>> = {};
    class ExcelError extends Error {
      constructor(
        readonly code: string,
        message?: string,
      ) {
        super(message);
      }
    }
    (globalThis as Record<string, unknown>).CustomFunctions = {
      associate: (id: string, fn: (...a: unknown[]) => Promise<unknown>) => (fns[id] = fn),
      Error: ExcelError,
      ErrorCode: { notAvailable: "#N/A", invalidValue: "#VALUE!" },
    };
    try {
      expect(registerCellFunctions(Promise.resolve(f.backend), true)).toBeDefined();
      expect(Object.keys(fns).sort()).toEqual(["ASK", "MODELS"]);
      expect(await fns.ASK("p", [["x"]], null, null, {})).toBe("svar på x");
      const err = await fns.ASK("", null, null, null, {}).catch((e) => e);
      expect(err).toBeInstanceOf(ExcelError);
      expect(err).toMatchObject({ code: "#VALUE!", message: "Prompten mangler" });
    } finally {
      delete (globalThis as Record<string, unknown>).CustomFunctions;
    }
  });
});
