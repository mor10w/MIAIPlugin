import { Endpoint, type TaskEvent } from "@miai/shared";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { DirectBackend } from "./direct";
import { addressSpace, localNetworkAccess } from "./localNetwork";

const HOSTED = "https://example.github.io/miai/taskpane.html";
const OLLAMA = "http://127.0.0.1:11434/v1";

/** Mocker Permissions API'et. `states` er svaret pr. tilladelsesnavn; et ukendt navn afvises som i browseren. */
function mockPermissions(states: Record<string, string>) {
  const query = vi.fn(async ({ name }: { name: string }) => {
    if (!(name in states)) throw new TypeError(`'${name}' is not a valid enum value of type PermissionName.`);
    return { state: states[name] };
  });
  vi.stubGlobal("navigator", { permissions: { query } });
  return query;
}

/** Task pane som i Office på nettet: hostet offentligt og vist i en ramme. */
function mockPage(opts: { href?: string; framed?: boolean; featurePolicy?: object } = {}) {
  const self = {};
  const href = opts.href ?? HOSTED;
  vi.stubGlobal("window", { self, top: opts.framed ? {} : self, location: { href, origin: new URL(href).origin } });
  vi.stubGlobal("document", { featurePolicy: opts.featurePolicy });
}

beforeEach(() => mockPage());
afterEach(() => vi.unstubAllGlobals());

describe("addressSpace", () => {
  it("genkender maskinen selv og det lokale net ud fra adressen", () => {
    for (const url of ["http://127.0.0.1:11434/v1", "http://localhost:1234/v1", "http://[::1]:8080/v1", "https://app.localhost/"])
      expect(addressSpace(url), url).toBe("loopback");
    for (const url of ["http://192.168.1.20:8000/v1", "http://10.0.0.5/v1", "http://172.20.1.1/v1", "https://llm.firma.local/v1", "http://[fd12::1]/v1"])
      expect(addressSpace(url), url).toBe("local");
  });

  it("andre adresser er ukendte (kan være på internettet)", () => {
    for (const url of ["https://api.example.com/v1", "http://172.32.0.1/v1", "http://8.8.8.8/v1", "https://example.github.io/", "ikke en url"])
      expect(addressSpace(url), url).toBeUndefined();
  });
});

describe("localNetworkAccess", () => {
  it("afvist tilladelse: denied – og der spørges om tilladelsen, der passer til adressen", async () => {
    const query = mockPermissions({ "loopback-network": "denied", "local-network": "granted" });
    expect(await localNetworkAccess(OLLAMA)).toBe("denied");
    expect(query).toHaveBeenLastCalledWith({ name: "loopback-network" });
    expect(await localNetworkAccess("http://192.168.1.20:8000/v1")).toBeUndefined();
    expect(query).toHaveBeenLastCalledWith({ name: "local-network" });
  });

  it("givet tilladelse: intet i vejen", async () => {
    mockPermissions({ "loopback-network": "granted" });
    expect(await localNetworkAccess(OLLAMA)).toBeUndefined();
  });

  it("API'et findes ikke, eller browseren kender ikke navnet: ukendt, som i ældre browsere og Office's webviews", async () => {
    vi.stubGlobal("navigator", {});
    expect(await localNetworkAccess(OLLAMA)).toBeUndefined();
    // Kun det samlede navn fra de første versioner findes: det nye navn afvises, og det gamle spørges der ikke om.
    const query = mockPermissions({ "local-network-access": "denied" });
    expect(await localNetworkAccess(OLLAMA)).toBeUndefined();
    expect(query).toHaveBeenCalledTimes(1);
    expect(query).not.toHaveBeenCalledWith({ name: "local-network-access" });
  });

  it("en adresse på internettet spørger ikke om tilladelsen", async () => {
    const query = mockPermissions({ "loopback-network": "denied", "local-network": "denied" });
    expect(await localNetworkAccess("https://api.example.com/v1")).toBeUndefined();
    expect(query).not.toHaveBeenCalled();
  });

  it("ubesvaret tilladelse meldes kun, når kravet gælder for siden", async () => {
    mockPermissions({ "loopback-network": "prompt", "local-network": "prompt" });
    expect(await localNetworkAccess(OLLAMA, HOSTED)).toBe("prompt");
    // En side på maskinen selv (bridge, udvikling) må kalde lokale adresser uden tilladelse.
    expect(await localNetworkAccess(OLLAMA, "https://localhost:3000/taskpane.html")).toBeUndefined();
    // En side på det lokale net skal have lov til maskinen selv, men ikke til det lokale net.
    expect(await localNetworkAccess(OLLAMA, "https://miai.firma.local/taskpane.html")).toBe("prompt");
    expect(await localNetworkAccess("https://llm.firma.local/v1", "https://miai.firma.local/taskpane.html")).toBeUndefined();
  });

  it("frame: siden uden om add-in'et giver ikke tilladelsen videre", async () => {
    mockPermissions({ "loopback-network": "denied" });
    const featurePolicy = { features: () => ["loopback-network", "local-network"], allowsFeature: () => false };
    mockPage({ framed: true, featurePolicy });
    expect(await localNetworkAccess(OLLAMA)).toBe("frame");
    // Øverste side, eller en browser, der ikke kender funktionen: brugerens eget nej.
    mockPage({ framed: false, featurePolicy });
    expect(await localNetworkAccess(OLLAMA)).toBe("denied");
    mockPage({ framed: true, featurePolicy: { features: () => ["camera"], allowsFeature: () => false } });
    expect(await localNetworkAccess(OLLAMA)).toBe("denied");
    mockPage({ framed: true, featurePolicy: { features: () => ["loopback-network"], allowsFeature: () => true } });
    expect(await localNetworkAccess(OLLAMA)).toBe("denied");
  });
});

describe("DirectBackend med blokeret kald (mocket fetch-fejl)", () => {
  const ollama = Endpoint.parse({ id: "ollama", name: "Ollama", baseUrl: OLLAMA });
  /** Sådan fejler fetch i browseren, både når porten er lukket, og når browseren blokerer kaldet. */
  const failingFetch = () => vi.fn(async () => Promise.reject(new TypeError("Failed to fetch")));

  it("afvist tilladelse: probe melder blocked, og der prøves ikke med no-cors", async () => {
    mockPermissions({ "loopback-network": "denied" });
    const fetch = failingFetch();
    vi.stubGlobal("fetch", fetch);
    expect(await new DirectBackend().probe(ollama)).toMatchObject({ ok: false, reason: "blocked", lna: "denied" });
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it("givet tilladelse: samme fejl er absent som før", async () => {
    mockPermissions({ "loopback-network": "granted" });
    vi.stubGlobal("fetch", failingFetch());
    const r = await new DirectBackend().probe(ollama);
    expect(r).toMatchObject({ ok: false, reason: "absent" });
    expect(r).not.toHaveProperty("lna");
  });

  it("API'et findes ikke: absent, og cors når no-cors-kaldet lykkes – som før", async () => {
    vi.stubGlobal("navigator", {});
    vi.stubGlobal("fetch", failingFetch());
    expect(await new DirectBackend().probe(ollama)).toMatchObject({ ok: false, reason: "absent" });
    const noCorsOk = vi.fn(async (_url: unknown, init?: RequestInit) => {
      if (init?.mode === "no-cors") return new Response(null);
      throw new TypeError("Failed to fetch");
    });
    vi.stubGlobal("fetch", noCorsOk);
    expect(await new DirectBackend().probe(ollama)).toMatchObject({ ok: false, reason: "cors" });
  });

  it("statustjek og opgaver melder blokeringen i stedet for en forbindelsesfejl med CORS-vejledning", async () => {
    mockPermissions({ "loopback-network": "denied" });
    vi.stubGlobal("fetch", failingFetch());
    const backend = new DirectBackend();
    await backend.saveSettings((await backend.settings()).settings); // standard: Ollama på 127.0.0.1:11434
    const health = await backend.health();
    expect(health.endpoint).toMatchObject({ reachable: false, blocked: "denied" });
    expect(health.endpoint.error).not.toContain("CORS");

    const events: TaskEvent[] = [];
    await backend.runTask("simplify", { context: { host: "word", text: "Tekst" }, params: {} }, (e) => events.push(e), new AbortController().signal);
    expect(events).toEqual([{ type: "error", code: "LNA_BLOCKED", message: expect.stringContaining("127.0.0.1:11434") }]);
  });

  it("uden afvist tilladelse er opgavens fejl den sædvanlige forbindelsesfejl", async () => {
    mockPermissions({ "loopback-network": "prompt" });
    vi.stubGlobal("fetch", failingFetch());
    const backend = new DirectBackend();
    const events: TaskEvent[] = [];
    await backend.runTask("simplify", { context: { host: "word", text: "Tekst" }, params: {} }, (e) => events.push(e), new AbortController().signal);
    expect(events).toMatchObject([{ type: "error", code: "UPSTREAM_UNREACHABLE" }]);
    expect((await backend.health()).endpoint).not.toHaveProperty("blocked");
  });
});
