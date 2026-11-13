import { describe, expect, it } from "vitest";
import { clearChat, documentKey, loadChat, MAX_STORED_DOCUMENTS, MAX_STORED_MESSAGES, saveChat, type StoredMessage } from "./chatStore";

function memory() {
  const data = new Map<string, string>();
  return {
    data,
    getItem: (k: string) => data.get(k) ?? null,
    setItem: (k: string, v: string) => void data.set(k, v),
    removeItem: (k: string) => void data.delete(k),
  };
}
const msg = (content: string, role: "user" | "assistant" = "user"): StoredMessage => ({ role, content });

describe("chatStore (#26)", () => {
  it("nøglen er stabil pr. dokument, og dokumenter uden adresse huskes ikke", () => {
    const a = documentKey("https://firma.sharepoint.com/Budget 2026.xlsx");
    expect(a).toBe(documentKey("https://firma.sharepoint.com/Budget 2026.xlsx"));
    expect(a).not.toBe(documentKey("https://firma.sharepoint.com/Budget 2027.xlsx"));
    expect(a).toMatch(/^[0-9a-f]{9,}$/);
    expect(documentKey("")).toBeNull();
    const store = memory();
    saveChat(null, [msg("hej")], store);
    expect(store.data.size).toBe(0);
    expect(loadChat(null, store)).toEqual([]);
  });

  it("gemmer og henter samtalen pr. dokument, og 'ny samtale' rydder den", () => {
    const store = memory();
    saveChat("a", [msg("Hvad er summen?"), { role: "assistant", content: "Sorterer.", model: "qwen3", planned: true, log: "- Sortér: done" }], store);
    saveChat("b", [msg("andet dokument")], store);
    expect(loadChat("a", store)).toEqual([
      msg("Hvad er summen?"),
      { role: "assistant", content: "Sorterer.", model: "qwen3", planned: true, log: "- Sortér: done" },
    ]);
    expect(loadChat("b", store)).toEqual([msg("andet dokument")]);
    clearChat("a", store);
    expect(loadChat("a", store)).toEqual([]);
    expect(loadChat("b", store)).toHaveLength(1);
  });

  it("holder sig inden for grænserne: de nyeste beskeder og de senest brugte dokumenter", () => {
    const store = memory();
    saveChat(
      "a",
      Array.from({ length: MAX_STORED_MESSAGES + 5 }, (_, i) => msg(`besked ${i}`)),
      store,
    );
    const kept = loadChat("a", store);
    expect(kept).toHaveLength(MAX_STORED_MESSAGES);
    expect(kept.at(-1)!.content).toBe(`besked ${MAX_STORED_MESSAGES + 4}`);
    // Én meget lang besked gemmes stadig (den nyeste), men ikke de ældre oven i.
    saveChat("lang", [msg("gammel"), msg("x".repeat(70_000))], store);
    expect(loadChat("lang", store)).toHaveLength(1);
    for (let i = 0; i < MAX_STORED_DOCUMENTS + 2; i++) saveChat(`doc${i}`, [msg("hej")], store);
    expect(loadChat("a", store)).toEqual([]);
    expect(loadChat(`doc${MAX_STORED_DOCUMENTS + 1}`, store)).toHaveLength(1);
  });

  it("ødelagt indhold i lageret giver en tom samtale i stedet for en fejl", () => {
    const store = memory();
    store.setItem("miai.chat.v1:a", "{ikke json");
    expect(loadChat("a", store)).toEqual([]);
    store.setItem("miai.chat.v1:a", JSON.stringify([{ role: "system", content: "x" }, { role: "user", content: 5 }, msg("ok")]));
    expect(loadChat("a", store)).toEqual([msg("ok")]);
  });
});
