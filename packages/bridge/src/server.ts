import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join, resolve, sep } from "node:path";
import Fastify, { type FastifyInstance, type FastifyServerOptions } from "fastify";
import fastifyStatic from "@fastify/static";
import {
  cachedModels,
  Endpoint,
  probeEndpoint,
  isUrlAllowed,
  Policy,
  Settings,
  TaskRequest,
  toTaskView,
  UserTaskFile,
  validateInput,
  type HealthInfo,
  type ModelInfo,
  type SettingsView,
  type TaskEvent,
  type UpstreamError,
} from "@miai/shared";
import { z } from "zod";
import { appendAudit, auditFile, readAuditTail } from "./audit.js";
import { PolicyError, type ConfigStore } from "./config.js";
import { invalidateUserTasks, loadUserTasks, machineTasksDir, mergedCatalog, userTasksDir } from "./userTasks.js";
import { tokenMatches } from "./identity.js";
import { runTask } from "./runner.js";
import { pickLang, renderStatusPage } from "./statusPage.js";

export const VERSION = "0.1.0";
const KEY_MASK = "********";

export interface ServerOptions {
  store: ConfigStore;
  port: number;
  /** Mappe med add-in'ets byggede filer (packages/addin/dist). */
  addinDir?: string;
  /**
   * Adgangstoken (#11). Når det er sat, kræver alle /api/*-kald headeren X-MIAI-Token, så andre
   * lokale processer (fx andre brugere på en delt maskine) ikke kan bruge bridgen.
   */
  token?: string;
  /** TLS (cert/key). Udeladt i tests. */
  https?: { cert: Buffer; key: Buffer };
  fastify?: FastifyServerOptions;
}

export async function buildServer({ store, port, addinDir, https, token, fastify }: ServerOptions): Promise<FastifyInstance> {
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
      if (token && !tokenMatches(token, req.headers["x-miai-token"])) {
        return reply.code(401).send({ error: "Manglende eller forkert adgangstoken. Geninstallér MIAI i Office (manifestet indeholder tokenet).", code: "TOKEN" });
      }
    }
  });

  app.addHook("onSend", async (_req, reply) => {
    reply.header("x-content-type-options", "nosniff");
    reply.header("referrer-policy", "no-referrer");
  });

  // --- API ------------------------------------------------------------------

  /** Bruges af /api/health og statussiden. Endpoints uden for politikken kontaktes ikke. */
  const health = async (): Promise<HealthInfo> => {
    const ep = store.activeEndpoint();
    try {
      if (!isUrlAllowed(ep.baseUrl, store.getPolicy())) throw new Error("Endpoint er ikke tilladt af politikken.");
      await cachedModels(ep, AbortSignal.timeout(3_000));
      return { ok: true, version: VERSION, endpoint: { id: ep.id, baseUrl: ep.baseUrl, reachable: true } };
    } catch (err) {
      return { ok: true, version: VERSION, endpoint: { id: ep.id, baseUrl: ep.baseUrl, reachable: false, error: (err as Error).message } };
    }
  };

  app.get("/api/health", health);

  // Statusside (#15): "/" åbnes typisk i en almindelig browser; selve UI'et er /taskpane.html i Office.
  const addinBuilt = !!addinDir && existsSync(addinDir);
  app.get("/", async (req, reply) =>
    reply.type("text/html; charset=utf-8").header("cache-control", "no-store").send(
      renderStatusPage({
        version: VERSION,
        health: await health(),
        settings: store.getSettings(),
        policy: store.getPolicy(),
        policySource: store.policySource,
        addinBuilt,
        configFile: store.file,
        lang: pickLang(req.headers["accept-language"]),
      }),
    ),
  );

  app.get("/api/tasks", async () => {
    const cat = mergedCatalog(store.getSettings(), store.dir);
    return cat.defs.map((t) => toTaskView(t, cat.customIds.has(t.id)));
  });

  app.get("/api/models", async (req, reply) => {
    const q = z.object({ endpoint: z.string().optional() }).parse(req.query);
    const s = store.getSettings();
    const ep = s.endpoints.find((e) => e.id === (q.endpoint ?? s.activeEndpoint)) ?? store.activeEndpoint();
    if (!isUrlAllowed(ep.baseUrl, store.getPolicy())) return reply.code(403).send({ error: "Endpoint ikke tilladt af politikken" });
    let host = ep.baseUrl;
    try {
      host = new URL(ep.baseUrl).host;
    } catch {
      // Ugyldig URL: bruges kun i audit-posten.
    }
    const started = Date.now();
    try {
      const models = await cachedModels(ep, AbortSignal.timeout(5_000));
      if (s.auditLog) await appendAudit(auditFile(store.dir), { ts: new Date().toISOString(), kind: "models", host, status: "ok", durationMs: Date.now() - started });
      return models.map((id): ModelInfo => ({ id, endpoint: ep.id }));
    } catch (err) {
      if (s.auditLog) await appendAudit(auditFile(store.dir), { ts: new Date().toISOString(), kind: "models", host, status: "error" });
      return reply.code(502).send({ error: (err as Error).message, code: (err as UpstreamError).code });
    }
  });

  /** Tester en (evt. ikke gemt) endpoint uden CORS-begrænsninger (#4). Maskerede nøgler slås op. */
  app.post("/api/probe", async (req, reply) => {
    const body = z.object({ endpoint: Endpoint }).safeParse(req.body);
    if (!body.success) return reply.code(400).send({ error: z.prettifyError(body.error) });
    const ep = body.data.endpoint;
    if (ep.apiKey === KEY_MASK) ep.apiKey = store.getSettings().endpoints.find((e) => e.id === ep.id)?.apiKey ?? "";
    return probeEndpoint(ep, store.getPolicy(), { signal: AbortSignal.timeout(6_000) });
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

  /** Kør en opgave (indbygget eller brugerdefineret). Svarer altid med en SSE-strøm af TaskEvent. */
  app.post("/api/tasks/:id", async (req, reply) => {
    const { id } = req.params as { id: string };
    const task = mergedCatalog(store.getSettings(), store.dir).defs.find((t) => t.id === id);
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

  // --- Brugerdefinerede opgaver (F14) ----------------------------------------
  // Filerne ligger i bruger-mappen (indstillingen customTasksDir, standard <config>/tasks)
  // og evt. i maskin-mappen (MIAI_TASKS_DIR). Maskin-opgaver er read-only via UI.

  // Filstier sendes ikke til webviewet.
  app.get("/api/user-tasks", async () => [...loadUserTasks(store.getSettings(), store.dir).values()].map(({ spec, source }) => ({ spec, source })));

  app.put("/api/user-tasks/:id", async (req, reply) => {
    const parsed = UserTaskFile.safeParse(req.body);
    if (!parsed.success) return reply.code(400).send({ error: z.prettifyError(parsed.error) });
    if (parsed.data.id !== (req.params as { id: string }).id) return reply.code(400).send({ error: "Id i URL og indhold er forskellige." });
    const s = store.getSettings();
    if (loadUserTasks(s, store.dir).get(parsed.data.id)?.source === "machine") {
      return reply.code(409).send({ error: "Opgaven er defineret af din organisation og kan ikke ændres." });
    }
    const dir = userTasksDir(s, store.dir);
    mkdirSync(dir, { recursive: true });
    writeFileSync(join(dir, `${parsed.data.id}.json`), JSON.stringify(parsed.data, null, 2));
    invalidateUserTasks();
    return { ok: true };
  });

  app.delete("/api/user-tasks/:id", async (req, reply) => {
    const id = (req.params as { id: string }).id;
    const s = store.getSettings();
    const loaded = loadUserTasks(s, store.dir).get(id);
    if (!loaded) return reply.code(404).send({ error: "Opgaven findes ikke." });
    if (loaded.source === "machine") return reply.code(409).send({ error: "Opgaven er defineret af din organisation og kan ikke slettes." });
    // Slet den fil, opgaven faktisk blev læst fra (kan ligge i en undermappe eller have et andet navn),
    // men kun hvis den ligger i brugerens opgavemappe.
    const dir = resolve(userTasksDir(s, store.dir));
    const file = resolve(loaded.file);
    if (!file.startsWith(dir + sep)) return reply.code(409).send({ error: "Opgavefilen ligger uden for opgavemappen." });
    rmSync(file);
    invalidateUserTasks();
    return { ok: true };
  });

  // --- Audit-log (kun metadata) ----------------------------------------------
  app.get("/api/audit", async () => readAuditTail(auditFile(store.dir), 100));

  // --- Add-in'ets statiske filer (same origin som API'et) -------------------
  // index: false – "/" er statussiden; manifestet peger direkte på /taskpane.html.
  if (addinBuilt) await app.register(fastifyStatic, { root: addinDir!, index: false });

  return app;
}
