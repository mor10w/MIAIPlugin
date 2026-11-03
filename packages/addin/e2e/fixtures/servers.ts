import { execFileSync } from "node:child_process";
import { createReadStream, existsSync, mkdtempSync, readFileSync, statSync } from "node:fs";
import { createServer as createHttp, type IncomingMessage, type Server } from "node:http";
import { createServer as createHttps } from "node:https";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export const DIST = resolve(fileURLToPath(new URL("../../dist", import.meta.url)));

/** Svar fra mock-modellen pr. forespørgsel. */
export type ModelReply =
  /** Streamet svar (SSE), ét chunk pr. element. promptTokens sendes som usage (kun med stream_options). */
  | { stream: string[]; promptTokens?: number; finishReason?: string }
  /** HTTP-fejl fra endpointen. */
  | { status: number; error: string }
  /** Server, der ignorerer stream=true og svarer med ét JSON-objekt. */
  | { json: string }
  /** Værktøjskald (OpenAI tool_calls), evt. med indledende tekst før kaldet. */
  | { toolCall: { name: string; args: unknown }; text?: string };

export interface ChatRequestBody {
  model: string;
  messages: { role: string; content: string; tool_calls?: unknown[] }[];
  stream?: boolean;
  presence_penalty?: number;
  response_format?: { type: string };
  tools?: unknown[];
  [key: string]: unknown;
}

/** OpenAI-kompatibel mock (som Ollama/LM Studio) på en tilfældig port. */
export class MockModel {
  readonly requests: ChatRequestBody[] = [];
  /** Bestemmer svaret ud fra forespørgslen. Standard: et kort streamet svar. */
  reply: (body: ChatRequestBody) => ModelReply = () => ({ stream: ["ok"] });
  private server?: Server;
  url = "";

  constructor(readonly opts: { cors?: boolean; models?: string[]; port?: number } = {}) {}

  get lastRequest(): ChatRequestBody | undefined {
    return this.requests.at(-1);
  }

  async start(): Promise<this> {
    this.server = createHttp(async (req, res) => {
      if (this.opts.cors !== false) {
        res.setHeader("access-control-allow-origin", "*");
        res.setHeader("access-control-allow-headers", "content-type, authorization");
      }
      if (req.method === "OPTIONS") return res.end();
      if (req.url === "/v1/models") {
        res.setHeader("content-type", "application/json");
        return res.end(JSON.stringify({ object: "list", data: (this.opts.models ?? ["qwen3:4b"]).map((id) => ({ id })) }));
      }
      if (req.url !== "/v1/chat/completions") {
        res.statusCode = 404;
        return res.end();
      }
      const body = JSON.parse(await readBody(req)) as ChatRequestBody;
      this.requests.push(body);
      const r = this.reply(body);
      if ("status" in r) {
        res.statusCode = r.status;
        res.setHeader("content-type", "application/json");
        return res.end(JSON.stringify({ error: { message: r.error } }));
      }
      if ("json" in r) {
        res.setHeader("content-type", "application/json");
        return res.end(JSON.stringify({ choices: [{ message: { content: r.json }, finish_reason: "stop" }] }));
      }
      res.writeHead(200, { "content-type": "text/event-stream" });
      const chunks =
        "toolCall" in r
          ? [
              ...(r.text ? [{ choices: [{ delta: { content: r.text } }] }] : []),
              {
                choices: [
                  {
                    delta: {
                      tool_calls: [
                        { index: 0, id: "call_1", type: "function", function: { name: r.toolCall.name, arguments: JSON.stringify(r.toolCall.args) } },
                      ],
                    },
                    finish_reason: "tool_calls",
                  },
                ],
              },
            ]
          : [
              ...r.stream.map((content, i) => ({
                choices: [{ delta: { content }, finish_reason: i === r.stream.length - 1 ? (r.finishReason ?? "stop") : null }],
              })),
              ...(r.promptTokens !== undefined && body.stream_options
                ? [{ choices: [], usage: { prompt_tokens: r.promptTokens, completion_tokens: 10 } }]
                : []),
            ];
      for (const c of chunks) {
        res.write(`data: ${JSON.stringify(c)}\n\n`);
        await new Promise((ok) => setTimeout(ok, 5)); // realistisk streaming: flere renders i UI'et
      }
      res.end("data: [DONE]\n\n");
    });
    this.url = `http://127.0.0.1:${await listen(this.server, this.opts.port)}/v1`;
    return this;
  }

  async stop(): Promise<void> {
    await close(this.server);
  }
}

/** Selvsigneret certifikat for localhost (Office kræver HTTPS). Laves én gang pr. testkørsel. */
export function testCert(): { cert: Buffer; key: Buffer } {
  const dir = process.env.MIAI_E2E_CERT_DIR ?? mkdtempSync(join(tmpdir(), "miai-e2e-cert-"));
  process.env.MIAI_E2E_CERT_DIR = dir;
  const cert = join(dir, "localhost.crt");
  const key = join(dir, "localhost.key");
  if (!existsSync(cert)) {
    execFileSync(
      "openssl",
      [
        "req",
        "-x509",
        "-newkey",
        "rsa:2048",
        "-nodes",
        "-days",
        "2",
        "-sha256",
        "-subj",
        "/CN=localhost",
        "-addext",
        "subjectAltName=DNS:localhost,IP:127.0.0.1",
        "-keyout",
        key,
        "-out",
        cert,
      ],
      { stdio: "ignore" },
    );
  }
  return { cert: readFileSync(cert), key: readFileSync(key) };
}

const TYPES: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json",
};

/** Hoster add-in'ets build statisk over HTTPS – som GitHub Pages (direkte tilstand). */
export async function startStaticHost(): Promise<{ origin: string; stop: () => Promise<void> }> {
  const server = createHttps(testCert(), (req, res) => {
    const path = normalize(decodeURIComponent((req.url ?? "/").split("?")[0]));
    const file = join(DIST, path);
    if (!file.startsWith(DIST) || !existsSync(file) || statSync(file).isDirectory()) {
      res.statusCode = 404;
      return res.end();
    }
    res.setHeader("content-type", TYPES[extname(file)] ?? "application/octet-stream");
    createReadStream(file).pipe(res);
  });
  const port = await listen(server);
  return { origin: `https://localhost:${port}`, stop: () => close(server) };
}

/** Starter den rigtige MIAI Bridge (bygget) med en midlertidig konfiguration, der peger på mock-modellen. */
export async function startBridge(
  modelUrl: string,
  settings: Record<string, unknown> = {},
): Promise<{ origin: string; token: string; stop: () => Promise<void> }> {
  const { ConfigStore } = await import("@miai/bridge/dist/config.js");
  // Testene må ikke skrive i udviklerens rigtige nøglering (#16).
  const { MemorySecrets } = await import("@miai/bridge/dist/secrets.js");
  const { buildServer } = await import("@miai/bridge/dist/server.js");
  const { Settings } = await import("@miai/shared");
  const dir = mkdtempSync(join(tmpdir(), "miai-e2e-bridge-"));
  const store = new ConfigStore(dir, join(dir, "ingen-politik.json"), new MemorySecrets());
  store.update(Settings.parse({ endpoints: [{ id: "mock", name: "Mock", baseUrl: modelUrl }], activeEndpoint: "mock", ...settings }));
  // Porten skal kendes, før serveren bygges (Host-tjekket), så find en ledig først.
  const port = await freePort();
  // forceCloseConnections: luk browserens keep-alive-forbindelser ved oprydning (ellers venter close ~15 s).
  // Adgangstoken som i det personlige manifest (#11).
  const token = `e2e-${Math.random().toString(36).slice(2)}${Math.random().toString(36).slice(2)}`;
  const app = await buildServer({ store, port, token, addinDir: DIST, https: testCert(), fastify: { forceCloseConnections: true } });
  await app.listen({ host: "127.0.0.1", port });
  return { origin: `https://localhost:${port}`, token, stop: () => app.close() };
}

function readBody(req: IncomingMessage): Promise<string> {
  return new Promise((ok, fail) => {
    let s = "";
    req.on("data", (c) => (s += c));
    req.on("end", () => ok(s));
    req.on("error", fail);
  });
}

function listen(server: Server | ReturnType<typeof createHttps>, port = 0): Promise<number> {
  return new Promise((ok, fail) => {
    server.once("error", fail);
    server.listen(port, "127.0.0.1", () => ok((server.address() as AddressInfo).port));
  });
}

/** Er en fast port ledig? (Tests, der bruger kendte værktøjsporte, springes over, hvis fx Ollama kører lokalt.) */
export async function portIsFree(port: number): Promise<boolean> {
  const s = createHttp();
  try {
    await listen(s, port);
    return true;
  } catch {
    return false;
  } finally {
    await close(s);
  }
}

function close(server?: { close: (cb: () => void) => void; closeAllConnections?: () => void }): Promise<void> {
  if (!server) return Promise.resolve();
  server.closeAllConnections?.();
  return new Promise((ok) => server.close(() => ok()));
}

async function freePort(): Promise<number> {
  const s = createHttp();
  const port = await listen(s);
  await close(s);
  return port;
}
