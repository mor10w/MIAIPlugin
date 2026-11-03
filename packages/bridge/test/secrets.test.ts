import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Settings } from "@miai/shared";
import { ConfigStore, KEYRING_REF } from "../src/config.js";
import { defaultSecretStore, FileSecrets, MemorySecrets, type SecretStore } from "../src/secrets.js";

/** API-nøgler i OS'ets nøglering (#16). Nøgleringen mockes med MemorySecrets. */

const tmp = () => mkdtempSync(join(tmpdir(), "miai-secrets-"));
const noPolicy = (dir: string) => join(dir, "ingen-politik.json");
const raw = (dir: string) => readFileSync(join(dir, "config.json"), "utf8");
const settings = (endpoints: { id: string; apiKey?: string }[]) =>
  Settings.parse({
    endpoints: endpoints.map((e) => ({ name: e.id, baseUrl: "http://localhost:11434/v1", ...e })),
    activeEndpoint: endpoints[0].id,
  });

afterEach(() => vi.restoreAllMocks());

describe("nøglering", () => {
  it("nye nøgler gemmes i nøgleringen; config.json har kun en reference", () => {
    const dir = tmp();
    const secrets = new MemorySecrets();
    const store = new ConfigStore(dir, noPolicy(dir), secrets);
    store.update(settings([{ id: "gw", apiKey: "sk-hemmelig" }]));

    expect(raw(dir)).not.toContain("sk-hemmelig");
    expect(raw(dir)).toContain(`"${KEYRING_REF}gw"`);
    expect(secrets.data.get("endpoint:gw")).toBe("sk-hemmelig");
    // I hukommelsen er nøglen stadig den rigtige, så kald til modellen virker.
    expect(store.getSettings().endpoints[0].apiKey).toBe("sk-hemmelig");
  });

  it("nøglen slås op i nøgleringen ved næste start", () => {
    const dir = tmp();
    const secrets = new MemorySecrets();
    new ConfigStore(dir, noPolicy(dir), secrets).update(settings([{ id: "gw", apiKey: "sk-1" }]));
    const again = new ConfigStore(dir, noPolicy(dir), secrets);
    expect(again.getSettings().endpoints[0].apiKey).toBe("sk-1");
    expect(again.keyStorage).toBe("keyring");
  });

  it("eksisterende nøgler i klartekst flyttes til nøgleringen ved første start", () => {
    const dir = tmp();
    // Som skrevet af en tidligere version.
    writeFileSync(join(dir, "config.json"), JSON.stringify({ settings: settings([{ id: "old", apiKey: "sk-gammel" }]) }));
    const secrets = new MemorySecrets();
    const store = new ConfigStore(dir, noPolicy(dir), secrets);

    expect(raw(dir)).not.toContain("sk-gammel");
    expect(secrets.data.get("endpoint:old")).toBe("sk-gammel");
    expect(store.getSettings().endpoints[0].apiKey).toBe("sk-gammel");
  });

  it("fjernede endpoints og slettede nøgler ryddes også op i nøgleringen", () => {
    const dir = tmp();
    const secrets = new MemorySecrets();
    const store = new ConfigStore(dir, noPolicy(dir), secrets);
    store.update(
      settings([
        { id: "a", apiKey: "sk-a" },
        { id: "b", apiKey: "sk-b" },
      ]),
    );
    store.update(settings([{ id: "a" }]));
    expect([...secrets.data.keys()]).toEqual([]);
  });

  it("endpoints uden nøgle skriver intet i nøgleringen", () => {
    const dir = tmp();
    const secrets = new MemorySecrets();
    new ConfigStore(dir, noPolicy(dir), secrets).update(settings([{ id: "ollama" }]));
    expect(secrets.data.size).toBe(0);
  });

  it("manglende nøgle i nøgleringen giver en tom nøgle og en advarsel – ikke en reference sendt til serveren", () => {
    const dir = tmp();
    const secrets = new MemorySecrets();
    new ConfigStore(dir, noPolicy(dir), secrets).update(settings([{ id: "gw", apiKey: "sk-1" }]));
    secrets.data.clear();
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const store = new ConfigStore(dir, noPolicy(dir), secrets);
    expect(store.getSettings().endpoints[0].apiKey).toBe("");
    expect(warn).toHaveBeenCalledWith(expect.stringContaining("findes ikke i nøgleringen"));
  });
});

describe("fallback uden nøglering", () => {
  it("nøgler bliver i config.json (0600), og bridgen melder 'file'", () => {
    const dir = tmp();
    const store = new ConfigStore(dir, noPolicy(dir), new FileSecrets());
    store.update(settings([{ id: "gw", apiKey: "sk-fil" }]));
    expect(raw(dir)).toContain("sk-fil");
    expect(store.keyStorage).toBe("file");
  });

  it("fejler nøgleringen ved gem, bliver nøglen i filen med en advarsel", () => {
    const dir = tmp();
    const broken: SecretStore = {
      kind: "keyring",
      get: () => null,
      set: () => {
        throw new Error("låst");
      },
      delete: () => {},
    };
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    new ConfigStore(dir, noPolicy(dir), broken).update(settings([{ id: "gw", apiKey: "sk-x" }]));
    expect(raw(dir)).toContain("sk-x");
    expect(warn).toHaveBeenCalledWith(expect.stringContaining("Kunne ikke gemme nøglen"));
  });

  it("MIAI_KEYRING=off slår nøgleringen fra", () => {
    expect(process.env.MIAI_KEYRING).toBe("off");
    expect(defaultSecretStore().kind).toBe("file");
  });

  it("utilgængelig nøglering giver fil-fallback med advarsel", () => {
    vi.stubEnv("MIAI_KEYRING", "");
    const warn = vi.fn();
    const store = defaultSecretStore(warn);
    // I CI/containere uden Secret Service er nøgleringen ikke tilgængelig; på en udviklermaskine kan den være det.
    if (store.kind === "file") expect(warn).toHaveBeenCalledWith(expect.stringContaining("nøglering kan ikke bruges"));
    else expect(warn).not.toHaveBeenCalled();
    vi.unstubAllEnvs();
  });
});
