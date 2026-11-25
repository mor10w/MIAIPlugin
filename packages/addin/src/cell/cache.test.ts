import { describe, expect, it, vi } from "vitest";
import { CELL_CACHE_KEY, CellCache, queryKey, type CacheStorage } from "./cache";

const memory = (): CacheStorage & { data: Map<string, string> } => {
  const data = new Map<string, string>();
  return { data, get: (k) => data.get(k) ?? null, set: (k, v) => void data.set(k, v), remove: (k) => void data.delete(k) };
};
const q = { endpoint: "http://127.0.0.1:11434/v1", model: "qwen3:4b", optimized: false, prompt: "Oversæt", input: "Hej" };

describe("cachenøglen", () => {
  it("er stabil og ændres af hvert felt", async () => {
    const base = await queryKey(q);
    expect(base).toMatch(/^[0-9a-f]{32}$/);
    expect(await queryKey({ ...q })).toBe(base);
    for (const other of [{ endpoint: "http://127.0.0.1:1234/v1" }, { model: "qwen3:32b" }, { optimized: true }, { prompt: "Opsummér" }, { input: "Hej." }])
      expect(await queryKey({ ...q, ...other })).not.toBe(base);
  });
  it("felterne kan ikke flyde sammen", async () => {
    expect(await queryKey({ ...q, prompt: "ab", input: "c" })).not.toBe(await queryKey({ ...q, prompt: "a", input: "bc" }));
  });
  it("indeholder hverken prompt eller input i klartekst", async () => {
    const key = await queryKey({ ...q, prompt: "hemmelig prompt", input: "fortrolig" });
    expect(key).not.toContain("hemmelig");
  });
});

describe("cachen", () => {
  it("overlever genstart og gemmer kun nøgle og svar", () => {
    vi.useFakeTimers();
    const store = memory();
    const a = new CellCache(store);
    a.set("k1", "Danmark");
    expect(store.data.has(CELL_CACHE_KEY)).toBe(false); // skrivningerne samles
    vi.advanceTimersByTime(600);
    expect(JSON.parse(store.data.get(CELL_CACHE_KEY)!)).toEqual([["k1", "Danmark"]]);
    expect(new CellCache(store).get("k1")).toBe("Danmark");
    vi.useRealTimers();
  });
  it("har et loft: det ældst brugte ryger først", () => {
    const c = new CellCache(memory(), { entries: 2, chars: 1_000 });
    c.set("a", "1");
    c.set("b", "2");
    expect(c.get("a")).toBe("1"); // a er nu senest brugt
    c.set("c", "3");
    expect([c.get("a"), c.get("b"), c.get("c")]).toEqual(["1", undefined, "3"]);
    const small = new CellCache(memory(), { entries: 100, chars: 10 });
    small.set("x", "123456");
    small.set("y", "123456");
    expect(small.size).toBe(1);
    small.set("z", "x".repeat(11)); // større end hele loftet: gemmes ikke
    expect(small.get("z")).toBeUndefined();
  });
  it("kan ryddes, og ødelagt indhold giver en tom cache", () => {
    const store = memory();
    const c = new CellCache(store);
    c.set("a", "1");
    c.flush();
    c.clear();
    expect(c.size).toBe(0);
    expect(store.data.has(CELL_CACHE_KEY)).toBe(false);
    store.data.set(CELL_CACHE_KEY, "{ikke json");
    expect(new CellCache(store).size).toBe(0);
  });
});
