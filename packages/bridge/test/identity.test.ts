import { mkdtempSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { createServer } from "node:net";
import { platform, tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { Settings } from "@miai/shared";
import { ConfigStore } from "../src/config.js";
import { DEFAULT_PORT, identityFile, loadIdentity, personalizeManifest, portIsFree, tokenMatches } from "../src/identity.js";
import { buildServer } from "../src/server.js";

const dir = () => mkdtempSync(join(tmpdir(), "miai-id-"));

describe("identitet pr. bruger (#11)", () => {
  it("opretter et token og husker det sammen med porten", async () => {
    const d = dir();
    const a = await loadIdentity(d, 45_101);
    const b = await loadIdentity(d, 45_101);
    expect(a.token).toHaveLength(43); // 32 tilfældige bytes, base64url
    expect(b).toEqual(a);
    if (platform() !== "win32") expect(statSync(identityFile(d)).mode & 0o777).toBe(0o600);
  });

  it("forskellige brugere (konfigurationsmapper) får forskellige tokens", async () => {
    expect((await loadIdentity(dir(), 45_102)).token).not.toBe((await loadIdentity(dir(), 45_102)).token);
  });

  const hold = (port: number) =>
    new Promise<ReturnType<typeof createServer>>((ok, fail) => {
      const srv = createServer();
      srv.once("error", fail).listen(port, "127.0.0.1", () => ok(srv));
    });
  const remember = (d: string, port: number) => writeFileSync(identityFile(d), JSON.stringify({ token: "t".repeat(43), port }));

  it("venter på den gemte port, når den kun er optaget kortvarigt (forrige udgave lukker)", async () => {
    const d = dir();
    remember(d, 45_210);
    const old = await hold(45_210);
    setTimeout(() => old.close(), 600);
    const started = Date.now();
    const id = await loadIdentity(d, undefined, { waitForStoredMs: 3_000 });
    expect(id.port).toBe(45_210); // samme port: manifestet i Office er stadig gyldigt
    expect(Date.now() - started).toBeGreaterThanOrEqual(500);
    expect(id.token).toBe("t".repeat(43));
  });

  it("vælger en anden port, når den gemte stadig er optaget efter ventetiden", async () => {
    const d = dir();
    remember(d, 45_211);
    const other = await hold(45_211);
    try {
      const id = await loadIdentity(d, undefined, { waitForStoredMs: 300 });
      expect(id.port).not.toBe(45_211);
      expect(JSON.parse(readFileSync(identityFile(d), "utf8")).port).toBe(id.port);
    } finally {
      other.close();
    }
  });

  it("vælger næste ledige port, når den gemte er optaget – og husker den", async () => {
    const d = dir();
    // Optag standardporten, som en anden brugers bridge ville gøre.
    const other = createServer();
    const blocked = await new Promise<boolean>((ok) => other.once("error", () => ok(false)).listen(DEFAULT_PORT, "127.0.0.1", () => ok(true)));
    try {
      const id = await loadIdentity(d);
      expect(id.port).not.toBe(DEFAULT_PORT);
      expect(id.port).toBeGreaterThan(DEFAULT_PORT);
      expect(await portIsFree(id.port)).toBe(true);
      expect(JSON.parse(readFileSync(identityFile(d), "utf8")).port).toBe(id.port);
      expect(blocked || !(await portIsFree(DEFAULT_PORT))).toBe(true);
    } finally {
      other.close();
    }
  });

  it("personligt manifest: port og token i task pane-adressen, ikke i ikoner", () => {
    const base = `<SourceLocation DefaultValue="https://localhost:44321/taskpane.html"/><bt:Url id="x" DefaultValue="https://localhost:44321/taskpane.html"/><bt:Image DefaultValue="https://localhost:44321/assets/icon-32.png"/><AppDomain>https://localhost:44321</AppDomain>`;
    const xml = personalizeManifest(base, { port: 44_322, token: "abc_DEF-123" });
    expect(xml.match(/taskpane\.html\?token=abc_DEF-123"/g)).toHaveLength(2);
    expect(xml).toContain('https://localhost:44322/assets/icon-32.png"');
    expect(xml).toContain("<AppDomain>https://localhost:44322</AppDomain>");
    expect(xml).not.toContain("44321");
  });

  it("personligt manifest med regnearksfunktioner (#80): tokenet føjes til ?functions=1, og funktionernes adresser følger porten", () => {
    const here = dirname(fileURLToPath(import.meta.url));
    const base = readFileSync(join(here, "../../addin/manifest.functions.xml"), "utf8");
    const xml = personalizeManifest(base, { port: 44_322, token: "abc_DEF-123" });
    // "&" skal være &amp; i XML; add-in'et læser både functions og token fra adressen.
    expect(xml).toContain('id="MIAI.Taskpane.Url" DefaultValue="https://localhost:44322/taskpane.html?functions=1&amp;token=abc_DEF-123"');
    expect(xml).toContain('<SourceLocation DefaultValue="https://localhost:44322/taskpane.html?token=abc_DEF-123"/>');
    expect(xml).toContain('DefaultValue="https://localhost:44322/functions.json"');
    expect(xml).toContain('DefaultValue="https://localhost:44322/functions-fallback.js"');
    expect(xml).not.toContain("44321");
    expect(xml).not.toMatch(/&(?!amp;)/); // ingen rå & i XML'en
    // Shared runtime kun i Excel: dér har knappen intet TaskpaneId; Word og PowerPoint har det stadig.
    expect(xml.match(/<Runtime /g)).toHaveLength(1);
    expect(xml.match(/<TaskpaneId>/g)).toHaveLength(2);
  });

  it("det udgivne bridge-manifest har ingen regnearksfunktioner (slås først til efter afprøvning i rigtig Excel)", () => {
    const here = dirname(fileURLToPath(import.meta.url));
    const xml = readFileSync(join(here, "../../addin/manifest.xml"), "utf8");
    expect(xml).not.toContain("CustomFunctions");
    expect(xml).not.toContain("<Runtimes>");
    expect(xml).not.toContain("functions=1");
    expect(xml.match(/<TaskpaneId>/g)).toHaveLength(3);
  });

  it("sammenligner tokens sikkert", () => {
    expect(tokenMatches("hemmelig", "hemmelig")).toBe(true);
    expect(tokenMatches("hemmelig", "hemmeliG")).toBe(false);
    expect(tokenMatches("hemmelig", "kort")).toBe(false);
    expect(tokenMatches("hemmelig", undefined)).toBe(false);
  });
});

describe("tokenkontrol i API'et", () => {
  const H = { host: "localhost:44321", "x-miai": "1" };
  async function app() {
    const d = dir();
    const store = new ConfigStore(d, join(d, "ingen-politik.json"));
    store.update(Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: "http://127.0.0.1:9/v1" }], activeEndpoint: "x" }));
    return buildServer({ store, port: 44321, token: "rigtigt-token" });
  }

  it("afviser API-kald uden eller med forkert token (401, kode TOKEN)", async () => {
    const a = await app();
    const none = await a.inject({ url: "/api/settings", headers: H });
    expect(none.statusCode).toBe(401);
    expect(none.json().code).toBe("TOKEN");
    expect((await a.inject({ url: "/api/settings", headers: { ...H, "x-miai-token": "forkert" } })).statusCode).toBe(401);
  });

  it("tillader API-kald med det rigtige token", async () => {
    const a = await app();
    expect((await a.inject({ url: "/api/settings", headers: { ...H, "x-miai-token": "rigtigt-token" } })).statusCode).toBe(200);
  });

  it("statussiden og statiske filer kræver intet token", async () => {
    const a = await app();
    expect((await a.inject({ url: "/", headers: { host: H.host } })).statusCode).toBe(200);
  });
});

describe("Outlook-manifest (#7)", () => {
  it("det personlige Outlook-manifest får port og adgangstoken overalt, hvor task pane indlæses", async () => {
    const { readFileSync, mkdtempSync } = await import("node:fs");
    const { tmpdir } = await import("node:os");
    const { join, resolve } = await import("node:path");
    const { writePersonalManifest } = await import("../src/identity.js");
    const base = resolve(__dirname, "../../addin/manifest.outlook.xml");
    const dir = mkdtempSync(join(tmpdir(), "miai-outlook-"));
    const out = writePersonalManifest(dir, base, { port: 45123, token: "hemmelig" }, "manifest-outlook.xml")!;
    expect(out).toBe(join(dir, "manifest-outlook.xml"));
    const xml = readFileSync(out, "utf8");
    expect(xml).toContain('xsi:type="MailApp"');
    expect(xml).not.toContain("44321");
    // ItemRead, ItemEdit og knappernes URL.
    expect(xml.match(/https:\/\/localhost:45123\/taskpane\.html\?token=hemmelig"/g)).toHaveLength(3);
    // Ikke samme Id som Word/Excel/PowerPoint-manifestet.
    const office = readFileSync(resolve(__dirname, "../../addin/manifest.xml"), "utf8");
    expect(xml.match(/<Id>([^<]+)<\/Id>/)![1]).not.toBe(office.match(/<Id>([^<]+)<\/Id>/)![1]);
  });
});
