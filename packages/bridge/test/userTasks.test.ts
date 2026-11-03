import { createServer, type IncomingMessage, type Server } from "node:http";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { FastifyInstance } from "fastify";
import { Settings, UserTaskFile, type TaskEvent } from "@miai/shared";
import { ConfigStore } from "../src/config.js";
import { buildServer } from "../src/server.js";

// Samme mock-server som bridge.test.ts: svarer på /v1/models og /v1/chat/completions.
let mock: Server;
let mockUrl: string;
let nextReply: string[] = [];
let lastBody: any;

async function readBody(req: IncomingMessage): Promise<string> {
  let s = "";
  for await (const c of req) s += c;
  return s;
}

beforeAll(async () => {
  mock = createServer(async (req, res) => {
    if (req.url === "/v1/models") {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ object: "list", data: [{ id: "mock-model" }] }));
    }
    if (req.url === "/v1/chat/completions") {
      lastBody = JSON.parse(await readBody(req));
      const parts = nextReply.shift()?.split("|") ?? ["ok"];
      res.writeHead(200, { "content-type": "text/event-stream" });
      for (const p of parts) res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: p } }] })}\n\n`);
      res.write(`data: ${JSON.stringify({ choices: [{ delta: {} }], usage: { prompt_tokens: 7, completion_tokens: 3 } })}\n\n`);
      return res.end("data: [DONE]\n\n");
    }
    res.statusCode = 404;
    res.end();
  });
  await new Promise<void>((r) => mock.listen(0, "127.0.0.1", r));
  const addr = mock.address() as { port: number };
  mockUrl = `http://127.0.0.1:${addr.port}/v1`;
});

afterAll(() => mock.close());

const PORT = 44321;
const H = { host: `localhost:${PORT}`, "x-miai": "1" };
let app: FastifyInstance;
let dir: string; // konfigurationsmappe (config.json, tasks/, audit.log)
let machineDir: string;

function events(payload: string): TaskEvent[] {
  return payload
    .split("\n\n")
    .filter((b) => b.startsWith("data: "))
    .map((b) => JSON.parse(b.slice(6)));
}

const spec = {
  id: "pris-tilbud",
  label: { da: "Pris-tilbud" },
  hosts: ["word"],
  input: "prompt",
  output: { format: "markdown" },
  params: [{ name: "kunde", label: { da: "Kunde" }, kind: "text" }],
  applyModes: ["insertAfter"],
  prompt: { system: "", user: "Tilbud for {{kunde}}:\n{{selection}}" },
};

beforeEach(async () => {
  dir = mkdtempSync(join(tmpdir(), "miai-ut-"));
  machineDir = mkdtempSync(join(tmpdir(), "miai-mt-"));
  process.env.MIAI_TASKS_DIR = machineDir;
  const store = new ConfigStore(dir, join(dir, "no-policy.json"));
  store.update(Settings.parse({ endpoints: [{ id: "mock", name: "Mock", baseUrl: mockUrl }], activeEndpoint: "mock" }));
  app = await buildServer({ store, port: PORT });
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
  rmSync(machineDir, { recursive: true, force: true });
});

function writeTask(file: string, s = spec) {
  mkdirSync(join(dir, "tasks"), { recursive: true });
  writeFileSync(file.endsWith(".json") ? file : join(dir, "tasks", `${file}.json`), JSON.stringify(s));
}

describe("brugerdefinerede opgaver", () => {
  it("læser opgave-filer fra bruger-mappen og markerer dem som custom", async () => {
    writeTask("pris-tilbud");
    const tasks = (await app.inject({ url: "/api/tasks", headers: H })).json() as any[];
    const t = tasks.find((x) => x.id === "pris-tilbud");
    expect(t).toMatchObject({ custom: true, hosts: ["word"], input: "prompt" });
    expect(tasks.find((x) => x.id === "grammar")).toMatchObject({ custom: false });
  });

  it("springer ugyldige filer over uden at fejle", async () => {
    writeTask("gyldig", { ...spec, id: "gyldig" });
    writeFileSync(join(dir, "tasks", "ugyldig.json"), "{ ikke json");
    const tasks = (await app.inject({ url: "/api/tasks", headers: H })).json() as any[];
    expect(tasks.some((x) => x.id === "gyldig")).toBe(true);
    expect(tasks.some((x) => x.id === "ugyldig")).toBe(false);
  });

  it("maskin-mappen vinder ved id-kollision, og UI kan ikke ændre den", async () => {
    writeTask("pris-tilbud");
    mkdirSync(machineDir, { recursive: true });
    writeFileSync(join(machineDir, "pris-tilbud.json"), JSON.stringify({ ...spec, label: { da: "Organisationens version" } }));
    const tasks = (await app.inject({ url: "/api/tasks", headers: H })).json() as any[];
    expect(tasks.find((x) => x.id === "pris-tilbud").label.da).toBe("Organisationens version");
    const listed = (await app.inject({ url: "/api/user-tasks", headers: H })).json() as any[];
    expect(listed.find((x) => x.spec.id === "pris-tilbud").source).toBe("machine");
    expect((await app.inject({ method: "PUT", url: "/api/user-tasks/pris-tilbud", headers: H, payload: spec })).statusCode).toBe(409);
    expect((await app.inject({ method: "DELETE", url: "/api/user-tasks/pris-tilbud", headers: H })).statusCode).toBe(409);
  });

  it("kører en brugerdefineret opgave og udfylder skabelonen", async () => {
    writeTask("pris-tilbud");
    nextReply = ["Tilbuddet"];
    const r = await app.inject({
      method: "POST",
      url: "/api/tasks/pris-tilbud",
      headers: H,
      payload: { context: { host: "word", text: "50 enheder" }, params: { kunde: "ACME", language: "Danish" } },
    });
    const ev = events(r.payload);
    expect(ev.filter((e) => e.type === "delta").length).toBeGreaterThan(0);
    expect(lastBody.messages[1].content).toContain("Tilbud for ACME:");
    expect(lastBody.messages[1].content).toContain("<document>\n50 enheder\n</document>");
  });

  it("PUT/DELETE skriver og fjerner fil i bruger-mappen", async () => {
    const file = join(dir, "tasks", "pris-tilbud.json");
    expect((await app.inject({ method: "PUT", url: "/api/user-tasks/pris-tilbud", headers: H, payload: spec })).statusCode).toBe(200);
    expect(existsSync(file)).toBe(true);
    expect(JSON.parse(readFileSync(file, "utf8")).label.da).toBe("Pris-tilbud");
    expect((await app.inject({ method: "DELETE", url: "/api/user-tasks/pris-tilbud", headers: H })).statusCode).toBe(200);
    expect(existsSync(file)).toBe(false);
  });

  it("GET /api/user-tasks returnerer tom liste uden filer", async () => {
    expect((await app.inject({ url: "/api/user-tasks", headers: H })).json()).toEqual([]);
  });

  it("PUT tillader overskrivning af et indbygget id (organisationsprompts)", async () => {
    const r = await app.inject({ method: "PUT", url: "/api/user-tasks/grammar", headers: H, payload: { ...spec, id: "grammar" } });
    expect(r.statusCode).toBe(200);
  });

  it("afviser ugyldig body ved PUT", async () => {
    const r = await app.inject({ method: "PUT", url: "/api/user-tasks/pris-tilbud", headers: H, payload: { ...spec, prompt: { user: "" } } });
    expect(r.statusCode).toBe(400);
  });

  it("UserTaskFile validerer filformatet", () => {
    expect(UserTaskFile.safeParse(spec).success).toBe(true);
  });
});

describe("audit-log", () => {
  it("logger opgavekørsel og model-henting (metadata, ikke indhold)", async () => {
    const store = new ConfigStore(dir, join(dir, "no-policy.json"));
    store.update(
      Settings.parse({ endpoints: [{ id: "mock", name: "Mock", baseUrl: mockUrl }], activeEndpoint: "mock", auditLog: true }),
    );
    app = await buildServer({ store, port: PORT });
    nextReply = ["ok"];
    const mockHost = new URL(mockUrl).host;

    await app.inject({ method: "GET", url: "/api/models", headers: H });
    await app.inject({ method: "POST", url: "/api/tasks/simplify", headers: H, payload: { context: { host: "word", text: "x" } } });

    const entries = (await app.inject({ url: "/api/audit", headers: H })).json() as any[];
    expect(entries).toHaveLength(2);
    expect(entries[0]).toMatchObject({ kind: "models", host: mockHost, status: "ok" });
    expect(entries[1]).toMatchObject({ kind: "task", task: "simplify", host: mockHost, model: "mock-model", status: "ok" });
    expect(entries[1].promptTokens).toBe(7);
    const raw = readFileSync(join(dir, "audit.log"), "utf8");
    expect(raw).not.toContain("x"); // aldrig indhold
  });

  it("logger fejl som status=error med kode", async () => {
    const dead = new ConfigStore(dir, join(dir, "no-policy.json"));
    dead.update(Settings.parse({ endpoints: [{ id: "dead", name: "Død", baseUrl: "http://127.0.0.1:9/v1" }], activeEndpoint: "dead", defaultModel: "m", auditLog: true }));
    const deadApp = await buildServer({ store: dead, port: PORT });
    const r = await deadApp.inject({ method: "POST", url: "/api/tasks/simplify", headers: H, payload: { context: { host: "word", text: "x" } } });
    expect(events(r.payload).at(-1)).toMatchObject({ type: "error" });
    const entries = (await deadApp.inject({ url: "/api/audit", headers: H })).json() as any[];
    expect(entries.at(-1)).toMatchObject({ kind: "task", status: "error" });
    expect(entries.at(-1).code).toBeTruthy();
  });

  it("logger intet, når auditLog er slået fra", async () => {
    const entries = (await app.inject({ url: "/api/audit", headers: H })).json() as any[];
    expect(entries).toEqual([]);
  });
});
