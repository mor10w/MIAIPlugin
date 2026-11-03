import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Settings } from "@miai/shared";
import { ConfigStore } from "../src/config.js";
import { buildServer } from "../src/server.js";
import { pickLang } from "../src/statusPage.js";

const PORT = 44321;
const host = `localhost:${PORT}`;

async function server(withAddin = true) {
  const dir = mkdtempSync(join(tmpdir(), "miai-status-"));
  const store = new ConfigStore(dir, join(dir, "ingen-politik.json"));
  // Ingen model kører på port 9 → "svarer ikke". Navnet tester escaping.
  store.update(Settings.parse({ endpoints: [{ id: "x", name: "<b>Min</b> Ollama", baseUrl: "http://127.0.0.1:9/v1", apiKey: "hemmelig-nøgle" }], activeEndpoint: "x" }));
  const addinDir = join(dir, "dist");
  if (withAddin) {
    const { mkdirSync } = await import("node:fs");
    mkdirSync(addinDir);
    writeFileSync(join(addinDir, "taskpane.html"), "<!doctype html><title>MIAI</title>task pane");
  }
  return buildServer({ store, port: PORT, addinDir: withAddin ? addinDir : undefined });
}

describe("statusside på /", () => {
  it("viser status på dansk uden X-MIAI-header", async () => {
    const app = await server();
    const r = await app.inject({ url: "/", headers: { host, "accept-language": "da-DK,da;q=0.9,en;q=0.8" } });
    expect(r.statusCode).toBe(200);
    expect(r.headers["content-type"]).toContain("text/html");
    expect(r.body).toContain("Bridgen kører");
    expect(r.body).toContain("Svarer ikke");
    expect(r.body).toContain("kører kun i Word, Excel og PowerPoint");
    expect(r.body).toContain("/taskpane.html");
  });

  it("viser engelsk ved andre sprog", async () => {
    const app = await server();
    const r = await app.inject({ url: "/", headers: { host, "accept-language": "en-GB,en" } });
    expect(r.body).toContain("The bridge is running");
  });

  it("lækker ikke API-nøgler og escaper navne", async () => {
    const app = await server();
    const r = await app.inject({ url: "/", headers: { host } });
    expect(r.body).not.toContain("hemmelig-nøgle");
    expect(r.body).toContain("&lt;b&gt;Min&lt;/b&gt; Ollama");
    expect(r.body).not.toContain("<b>Min</b>");
  });

  it("taskpane.html serveres stadig til Office", async () => {
    const app = await server();
    const r = await app.inject({ url: "/taskpane.html", headers: { host } });
    expect(r.statusCode).toBe(200);
    expect(r.body).toContain("task pane");
  });

  it("melder, hvis add-in'et ikke er bygget", async () => {
    const app = await server(false);
    const r = await app.inject({ url: "/", headers: { host, "accept-language": "da" } });
    expect(r.body).toContain("ikke bygget. Kør npm run build");
  });

  it("Host-tjekket gælder også statussiden", async () => {
    const app = await server();
    expect((await app.inject({ url: "/", headers: { host: "evil.example" } })).statusCode).toBe(421);
  });

  it("vælger sprog ud fra Accept-Language", () => {
    expect(pickLang("da")).toBe("da");
    expect(pickLang("nb-NO,en")).toBe("da");
    expect(pickLang("de-DE")).toBe("en");
    expect(pickLang(undefined)).toBe("en");
  });
});
