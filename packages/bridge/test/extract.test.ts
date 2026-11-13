import { mkdtempSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";
import type { FastifyInstance } from "fastify";
import { ConfigStore } from "../src/config.js";
import { buildServer } from "../src/server.js";

const PORT = 44321;
const H = { host: `localhost:${PORT}`, "x-miai": "1" };
let app: FastifyInstance;

beforeEach(async () => {
  const dir = mkdtempSync(join(tmpdir(), "miai-"));
  app = await buildServer({ store: new ConfigStore(dir, join(dir, "no-policy.json")), port: PORT });
});

const b64 = (s: string) => Buffer.from(s, "utf8").toString("base64");
const leftovers = () => readdirSync(tmpdir()).filter((f) => f.startsWith("miai-fil-"));

describe("udtræk af vedhæftede filer (#27)", () => {
  it("returnerer filens tekst og efterlader ingen midlertidig fil", async () => {
    const before = leftovers().length;
    const r = await app.inject({ method: "POST", url: "/api/extract", headers: H, payload: { name: "../../noter.md", data: b64("# Budget\nOmsætning: 760") } });
    expect(r.statusCode).toBe(200);
    expect(r.json()).toEqual({ text: "# Budget\nOmsætning: 760" });
    expect(leftovers().length).toBe(before);
  });

  it("afviser filtyper, der ikke kan læses, og melder ulæselige filer forståeligt", async () => {
    const exe = await app.inject({ method: "POST", url: "/api/extract", headers: H, payload: { name: "program.exe", data: b64("MZ") } });
    expect(exe.statusCode).toBe(400);
    expect(exe.json().error).toContain(".exe");
    const pdf = await app.inject({ method: "POST", url: "/api/extract", headers: H, payload: { name: "rapport.pdf", data: b64("ikke en pdf") } });
    expect(pdf.statusCode).toBe(422);
    expect(pdf.json().error).toContain("rapport.pdf");
    expect((await app.inject({ method: "POST", url: "/api/extract", headers: H, payload: { name: "x.txt" } })).statusCode).toBe(400);
  });

  it("opgaver og eksport tager imod mere end standardens 2 MB (billeder som data-URL, store ark)", async () => {
    const big = "A".repeat(3 * 1024 * 1024);
    const task = await app.inject({
      method: "POST",
      url: "/api/tasks/findes-ikke",
      headers: H,
      payload: { context: { host: "word", text: "" }, attachments: [{ kind: "image", name: "a.png", dataUrl: `data:image/png;base64,${big}` }] },
    });
    // 404 (ukendt opgave) – ikke 413: kroppen blev læst.
    expect(task.statusCode).toBe(404);
    const exp = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: { fileName: "x.xlsx", junk: big } });
    expect(exp.statusCode).toBe(400);
  });
});
