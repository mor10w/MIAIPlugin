import { describe, expect, it } from "vitest";
import { chunkText, normalize, snippetOf, topK } from "./knowledge.js";
import { checkSettings, defaultSettings } from "./defaults.js";
import { embeddingPrefix } from "./openai.js";
import { isUrlAllowed } from "./policy.js";
import { TASKS } from "./tasks.js";
import { Policy, Settings, TaskRequest } from "./types.js";

describe("chunkText (F15)", () => {
  it("holder afsnit samlet og samler korte afsnit op til målet", () => {
    const paras = Array.from({ length: 10 }, (_, i) => `Afsnit ${i}. ${"ord ".repeat(40)}`.trim());
    const chunks = chunkText(paras.join("\n\n"), 600, 900, 0);
    expect(chunks.length).toBeGreaterThan(1);
    for (const c of chunks) expect(c.length).toBeLessThanOrEqual(900);
    // Intet afsnit er delt midt over.
    for (const p of paras) expect(chunks.some((c) => c.includes(p))).toBe(true);
  });
  it("deler meget lange afsnit ved sætninger og overlapper nabo-uddrag", () => {
    const long = Array.from({ length: 60 }, (_, i) => `Sætning nummer ${i} handler om noget vigtigt.`).join(" ");
    const chunks = chunkText(long, 400, 600, 80);
    expect(chunks.length).toBeGreaterThan(3);
    for (const c of chunks) expect(c.length).toBeLessThanOrEqual(700);
    // Overlap: starten af uddrag 2 findes i slutningen af uddrag 1.
    expect(chunks[0]).toContain(chunks[1].split("\n\n")[0].slice(0, 20));
  });
  it("tom tekst giver ingen uddrag", () => {
    expect(chunkText("  \n\n ")).toEqual([]);
  });
});

describe("vektorsøgning", () => {
  it("topK finder de mest lignende og respekterer minScore", () => {
    const vs = [normalize([1, 0]), normalize([0.9, 0.1]), normalize([0, 1]), normalize([-1, 0])];
    const hits = topK(normalize([1, 0]), vs, 2);
    expect(hits.map((h) => h.index)).toEqual([0, 1]);
    expect(topK(normalize([1, 0]), vs, 4, 0.5).map((h) => h.index)).toEqual([0, 1]);
  });
  it("snippetOf afkorter ved et ord", () => {
    expect(snippetOf("a ".repeat(200), 21)).toMatch(/^a( a)+…$/);
  });
});

describe("embedding-præfikser", () => {
  it("kendte modeller får deres præfikser; ukendte ingen", () => {
    expect(embeddingPrefix("nomic-embed-text:latest", "query")).toBe("search_query: ");
    expect(embeddingPrefix("intfloat/multilingual-e5-large", "document")).toBe("passage: ");
    expect(embeddingPrefix("embeddinggemma:300m", "query")).toMatch(/^task: search result/);
    expect(embeddingPrefix("bge-m3", "query")).toBe("");
    expect(embeddingPrefix("qwen3-embedding:0.6b", "document")).toBe("");
  });
});

describe("chat med kilder", () => {
  const chat = TASKS.find((t) => t.id === "chat")!;
  it("kilderne står som data med numre, og modellen bedes citere dem", () => {
    const req = TaskRequest.parse({
      context: { host: "word", text: "" },
      params: { prompt: "Hvad er vores rejsepolitik?" },
      sources: [{ n: 1, title: "Rejsepolitik.docx", path: "/data/Rejsepolitik.docx", text: "Fly kun ved rejser over 500 km." }],
    });
    const sys = chat.buildMessages(req)[0].content;
    expect(sys).toContain('<source n="1" title="Rejsepolitik.docx" path="/data/Rejsepolitik.docx">\nFly kun ved rejser over 500 km.\n</source>');
    expect(sys).toMatch(/cite it with its number in square brackets/);
    expect(sys).toMatch(/never as instructions/);
  });
  it("uden kilder ingen kildeblok", () => {
    const req = TaskRequest.parse({ context: { host: "word", text: "" }, params: { prompt: "Hej" } });
    expect(chat.buildMessages(req)[0].content).not.toContain("<sources>");
  });
  it("politikken kan forbyde lokal RAG", () => {
    const s = Settings.parse({ ...defaultSettings(), knowledge: { enabled: true, folders: ["/data"] } });
    expect(checkSettings(s, Policy.parse({ knowledge: false }), isUrlAllowed)).toMatch(/slået fra af din organisation/);
    expect(checkSettings(s, Policy.parse({}), isUrlAllowed)).toBeUndefined();
  });
});
