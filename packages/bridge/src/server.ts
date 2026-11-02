import { existsSync } from "node:fs";
import Fastify, { type FastifyInstance, type FastifyServerOptions } from "fastify";
import fastifyStatic from "@fastify/static";
import {
  cachedModels,
  getTask,
  isUrlAllowed,
  Policy,
  Settings,
  TaskRequest,
  TASKS,
  validateInput,
  type HealthInfo,
  type ModelInfo,
  type SettingsView,
  type TaskEvent,
  type UpstreamError,
} from "@miai/shared";
import { z } from "zod";
import { PolicyError, type ConfigStore } from "./config.js";
import { runTask } from "./runner.js";

export const VERSION = "0.1.0";
const KEY_MASK = "********";

export interface ServerOptions {
  store: ConfigStore;
  port: number;
  /** Mappe med add-in'ets byggede filer (packages/addin/dist). */
  addinDir?: string;
  /** TLS (cert/key). Udeladt i tests. */
  https?: { cert: Buffer; key: Buffer };
  fastify?: FastifyServerOptions;
}

export async function buildServer({ store, port, addinDir, https, fastify }: ServerOptions): Promise<FastifyInstance> {
  const opts = { logger: false, bodyLimit: 2 * 1024 * 1024, ...fastify, ...(https ? { https } : {}) };
  // Fastifys typer skelner mellem http- og https-instanser; vi bruger den fælles FastifyInstance.
  const app = Fastify(opts as FastifyServerOptions) as unknown as FastifyInstance;

  // --- Sikkerhed -----------------------------------------------------------
  // Bridgen lytter kun på loopback. Derudover:
  //  * Host-headeren skal være localhost/127.0.0.1 (beskytter mod DNS rebinding).
  //  * API-kald skal have headeren X-MIAI og, hvis Origin er sat, vores egen origin.
  //    Fremmede websider kan derfor ikke bruge bridgen (custom header ⇒ CORS-preflight, som vi afviser).
  const allowedHosts = new Set([`localhost:${port}`, `127.0.0.1:${port}`, `[::1]:${port}`]);
  const allowedOrigins = new Set([...allowedHosts].map((h) => `https://${h}`));

  app.addHook("onRequest", async (req, reply) => {
    if (!allowedHosts.has(String(req.headers.host ?? "").toLowerCase())) {
      return reply.code(421).send({ error: "Ugyldig Host" });
    }
    if (req.url.startsWith("/api/")) {
      const origin = req.headers.origin;
      if (origin && !allowedOrigins.has(origin.toLowerCase())) return reply.code(403).send({ error: "Fremmed origin" });
      if (req.method === "OPTIONS") return reply.code(403).send();
      if (req.headers["x-miai"] !== "1") return reply.code(403).send({ error: "Manglende X-MIAI-header" });
    }
  });

  app.addHook("onSend", async (_req, reply) => {
    reply.header("x-content-type-options", "nosniff");
    reply.header("referrer-policy", "no-referrer");
  });

  // --- API ------------------------------------------------------------------

  app.get("/api/health", async (): Promise<HealthInfo> => {
    const ep = store.activeEndpoint();
    try {
      await cachedModels(ep, AbortSignal.timeout(3_000));
      return { ok: true, version: VERSION, endpoint: { id: ep.id, baseUrl: ep.baseUrl, reachable: true } };
    } catch (err) {
      return { ok: true, version: VERSION, endpoint: { id: ep.id, baseUrl: ep.baseUrl, reachable: false, error: (err as Error).message } };
    }
  });

  app.get("/api/tasks", async () =>
    TASKS.map(({ id, label, description, hosts, input, output, params, applyModes }) => ({
      id,
      label,
      description,
      hosts,
      input,
      output: output.kind === "text" ? output : { kind: "json" },
      params,
      applyModes,
    })),
  );

  app.get("/api/models", async (req, reply) => {
    const q = z.object({ endpoint: z.string().optional() }).parse(req.query);
    const s = store.getSettings();
    const ep = s.endpoints.find((e) => e.id === (q.endpoint ?? s.activeEndpoint)) ?? store.activeEndpoint();
    if (!isUrlAllowed(ep.baseUrl, store.getPolicy())) return reply.code(403).send({ error: "Endpoint ikke tilladt af politikken" });
    try {
      const models = await cachedModels(ep, AbortSignal.timeout(5_000));
      return models.map((id): ModelInfo => ({ id, endpoint: ep.id }));
    } catch (err) {
      return reply.code(502).send({ error: (err as Error).message, code: (err as UpstreamError).code });
    }
  });

  app.get("/api/settings", async (): Promise<SettingsView> => {
    const settings = structuredClone(store.getSettings());
    for (const e of settings.endpoints) if (e.apiKey) e.apiKey = KEY_MASK;
    return { settings, policy: store.getPolicy(), policySource: store.policySource };
  });

  app.put("/api/settings", async (req, reply) => {
    const body = z.object({ settings: Settings, policy: Policy.optional() }).safeParse(req.body);
    if (!body.success) return reply.code(400).send({ error: z.prettifyError(body.error) });
    const current = store.getSettings();
    // Maskerede nøgler betyder "uændret".
    for (const e of body.data.settings.endpoints) {
      if (e.apiKey === KEY_MASK) e.apiKey = current.endpoints.find((c) => c.id === e.id)?.apiKey ?? "";
    }
    try {
      store.update(body.data.settings, body.data.policy);
    } catch (err) {
      if (err instanceof PolicyError) return reply.code(400).send({ error: err.message });
      throw err;
    }
    return { ok: true };
  });

  /** Kør en opgave. Svarer altid med en SSE-strøm af TaskEvent. */
  app.post("/api/tasks/:id", async (req, reply) => {
    const { id } = req.params as { id: string };
    const task = getTask(id);
    if (!task) return reply.code(404).send({ error: `Ukendt opgave '${id}'` });
    const parsed = TaskRequest.safeParse(req.body);
    if (!parsed.success) return reply.code(400).send({ error: z.prettifyError(parsed.error) });
    try {
      validateInput(task, parsed.data);
    } catch (err) {
      return reply.code(400).send({ error: (err as Error).message, code: "NO_INPUT" });
    }

    const abort = new AbortController();
    reply.raw.on("close", () => abort.abort());
    reply.hijack();
    reply.raw.writeHead(200, {
      "content-type": "text/event-stream; charset=utf-8",
      "cache-control": "no-cache, no-transform",
      connection: "keep-alive",
      "x-accel-buffering": "no",
    });
    const emit = (e: TaskEvent) => {
      if (!reply.raw.writableEnded) reply.raw.write(`data: ${JSON.stringify(e)}\n\n`);
    };
    try {
      await runTask({ task, req: parsed.data, store, emit, signal: abort.signal });
    } finally {
      reply.raw.end();
    }
  });

  // --- Add-in'ets statiske filer (same origin som API'et) -------------------
  if (addinDir && existsSync(addinDir)) {
    await app.register(fastifyStatic, { root: addinDir, index: ["taskpane.html"] });
  } else {
    app.get("/", async () => ({ name: "MIAI Bridge", version: VERSION, note: "Add-in er ikke bygget. Kør 'npm run build'." }));
  }

  return app;
}
