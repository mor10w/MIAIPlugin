import { describe, expect, it } from "vitest";
import { buildCsp, cspHostSources } from "./csp";

describe("Content-Security-Policy (#28)", () => {
  it("politikkens hosts bliver til kilder for http og https på alle porte", () => {
    expect(cspHostSources(["127.0.0.1", "localhost", "::1"])).toEqual([
      "http://127.0.0.1:*",
      "https://127.0.0.1:*",
      "http://localhost:*",
      "https://localhost:*",
      "http://[::1]:*",
      "https://[::1]:*",
    ]);
    expect(cspHostSources(["*.Firma.local", " llm.firma.local "])).toEqual([
      "http://*.firma.local:*",
      "https://*.firma.local:*",
      "http://llm.firma.local:*",
      "https://llm.firma.local:*",
    ]);
    expect(cspHostSources(["*"])).toEqual(["http:", "https:"]);
  });

  it("ugyldige navne kan ikke smugle direktiver ind", () => {
    expect(cspHostSources(["evil.example; script-src *", "a b", "'unsafe-inline'", "http://x", ""])).toEqual([]);
  });

  it("direkte tilstand tillader politikkens hosts; med bridge kun add-in'ets egen adresse", () => {
    const direct = buildCsp("direct", { allowedHosts: ["127.0.0.1"] });
    expect(direct).toContain("connect-src 'self' http://127.0.0.1:* https://127.0.0.1:*;");
    expect(direct).toContain("img-src 'self' data: blob:");
    expect(direct).toContain("object-src 'none'");
    expect(direct).toContain("form-action 'none'");
    // Scripts begrænses bevidst ikke (Office.js og kodetilstandens sandkasse).
    expect(direct).not.toContain("script-src");
    expect(buildCsp("bridge", { allowedHosts: ["127.0.0.1", "llm.firma.local"] })).toContain("connect-src 'self';");
  });
});
