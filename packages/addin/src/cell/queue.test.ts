import { describe, expect, it } from "vitest";
import { CellCanceled, CellQueue, CellStopped } from "./queue";

/** Et stykke arbejde, testen selv afslutter. */
function deferredWork() {
  const calls: { key: string; signal: AbortSignal; done: (v: string) => void; fail: (e: unknown) => void }[] = [];
  const work = (key: string) => (signal: AbortSignal) =>
    new Promise<string>((done, fail) => {
      calls.push({ key, signal, done, fail });
      signal.addEventListener("abort", () => fail(new Error("aborted")));
    });
  return { calls, work };
}
const tick = () => new Promise((r) => setTimeout(r, 0));

describe("køen for MIAI.ASK", () => {
  it("respekterer loftet over samtidige kald og fortsætter, når et kald er færdigt", async () => {
    const { calls, work } = deferredWork();
    const q = new CellQueue(() => 2);
    const tickets = ["a", "b", "c", "d"].map((k) => q.enqueue(k, work(k)));
    expect(q.state).toEqual({ running: 2, queued: 2 });
    calls[0].done("A");
    expect(await tickets[0].promise).toBe("A");
    await tick();
    expect(calls.map((c) => c.key)).toEqual(["a", "b", "c"]);
    expect(q.state).toEqual({ running: 2, queued: 1 });
  });

  it("celler med samme forespørgsel deler ét kald", async () => {
    const { calls, work } = deferredWork();
    const q = new CellQueue(() => 2);
    const t1 = q.enqueue("ens", work("ens"));
    const t2 = q.enqueue("ens", work("ens"));
    expect(calls).toHaveLength(1);
    calls[0].done("svar");
    expect(await Promise.all([t1.promise, t2.promise])).toEqual(["svar", "svar"]);
  });

  it("afbrydelse fra Excel: ud af køen, eller afbryd kaldet – men først når ingen celler venter", async () => {
    const { calls, work } = deferredWork();
    const q = new CellQueue(() => 1);
    const run1 = q.enqueue("a", work("a"));
    const run2 = q.enqueue("a", work("a"));
    const waiting = q.enqueue("b", work("b"));
    waiting.cancel();
    await expect(waiting.promise).rejects.toBeInstanceOf(CellCanceled);
    expect(q.state).toEqual({ running: 1, queued: 0 });
    run1.cancel();
    expect(calls[0].signal.aborted).toBe(false); // den anden celle venter stadig
    run2.cancel();
    expect(calls[0].signal.aborted).toBe(true);
    await expect(run1.promise).rejects.toBeInstanceOf(CellCanceled);
    // Samme nøgle kan sættes i kø igen efter en afbrydelse.
    await tick();
    const again = q.enqueue("a", work("a"));
    await tick();
    calls.at(-1)!.done("nyt");
    expect(await again.promise).toBe("nyt");
  });

  it("Stop tømmer køen og afbryder kald i gang", async () => {
    const { calls, work } = deferredWork();
    const q = new CellQueue(() => 1);
    const seen: number[] = [];
    q.subscribe((s) => seen.push(s.running + s.queued));
    const tickets = ["a", "b", "c"].map((k) => q.enqueue(k, work(k)));
    expect(q.stopAll()).toBe(3);
    for (const t of tickets) await expect(t.promise).rejects.toBeInstanceOf(CellStopped);
    expect(calls).toHaveLength(1);
    expect(calls[0].signal.aborted).toBe(true);
    await tick();
    expect(q.state).toEqual({ running: 0, queued: 0 });
    expect(seen.at(-1)).toBe(0);
  });

  it("en fejl i ét kald stopper ikke de næste", async () => {
    const { calls, work } = deferredWork();
    const q = new CellQueue(() => 1);
    const bad = q.enqueue("a", work("a"));
    const good = q.enqueue("b", work("b"));
    calls[0].fail(new Error("intet net"));
    await expect(bad.promise).rejects.toThrow("intet net");
    await tick();
    calls[1].done("ok");
    expect(await good.promise).toBe("ok");
  });
});
