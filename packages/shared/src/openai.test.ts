import { createServer, type Server } from "node:http";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { streamChat, UpstreamError } from "./openai.js";
import { Endpoint } from "./types.js";

/** Mock-server, hvis opførsel styres af stien: /slow-stream, /silent-start, /stall. */
let server: Server;
let base = "";
const sse = (text: string) => `data: ${JSON.stringify({ choices: [{ delta: { content: text }, finish_reason: null }] })}\n\n`;
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

beforeAll(async () => {
  server = createServer(async (req, res) => {
    const mode = req.url!.split("/")[1];
    if (mode === "silent-start") await wait(1_500);
    res.writeHead(200, { "content-type": "text/event-stream" });
    if (mode === "silent-start") res.write(sse("for sent"));
    else {
      // Fem stykker med 400 ms imellem: 2 s i alt – længere end timeouten på 1 s.
      for (let i = 0; i < 5; i++) {
        res.write(sse(`del${i} `));
        if (mode === "stall" && i === 1) await wait(10_000);
        await wait(400);
      }
    }
    res.end("data: [DONE]\n\n");
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const addr = server.address() as { port: number };
  base = `http://127.0.0.1:${addr.port}`;
});
afterAll(() => {
  server.closeAllConnections();
  server.close();
});

const ep = (mode: string) => Endpoint.parse({ id: "t", name: "Test", baseUrl: `${base}/${mode}`, timeoutMs: 1_000 });
async function collect(mode: string): Promise<string> {
  let out = "";
  for await (const c of streamChat(ep(mode), { model: "m", messages: [{ role: "user", content: "hej" }] })) out += c.text ?? "";
  return out;
}

describe("timeout som tid uden svar", () => {
  it("afbryder ikke et svar, der strømmer jævnt, selvom det varer længere end timeouten", async () => {
    expect(await collect("slow-stream")).toBe("del0 del1 del2 del3 del4 ");
  });

  it("opgiver en server, der er tavs før svaret", async () => {
    await expect(collect("silent-start")).rejects.toMatchObject({ code: "UPSTREAM_TIMEOUT" });
  });

  it("opgiver et svar, der går i stå midtvejs, med en forståelig fejl", async () => {
    const err = await collect("stall").catch((e: unknown) => e);
    expect(err).toBeInstanceOf(UpstreamError);
    expect(err).toMatchObject({ code: "UPSTREAM_TIMEOUT" });
    expect((err as Error).message).toMatch(/Test har ikke svaret i 1 s/);
  });

  it("standarden er 5 minutter, og op til 30 minutter kan vælges", () => {
    expect(Endpoint.parse({ id: "a", name: "A", baseUrl: "http://x/v1" }).timeoutMs).toBe(300_000);
    expect(Endpoint.parse({ id: "a", name: "A", baseUrl: "http://x/v1", timeoutMs: 1_800_000 }).timeoutMs).toBe(1_800_000);
    expect(() => Endpoint.parse({ id: "a", name: "A", baseUrl: "http://x/v1", timeoutMs: 1_800_001 })).toThrow();
  });
});
