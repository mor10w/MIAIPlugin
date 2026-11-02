import { Policy, Settings } from "./types.js";

/** Standardindstillinger: Ollama på brugerens egen maskine. */
export function defaultSettings(): Settings {
  return Settings.parse({
    endpoints: [{ id: "ollama", name: "Ollama", baseUrl: "http://127.0.0.1:11434/v1", structuredOutput: "json_schema" }],
    activeEndpoint: "ollama",
  });
}

/** Standardpolitik: kun loopback. */
export function defaultPolicy(): Policy {
  return Policy.parse({});
}

export function activeEndpoint(s: Settings) {
  return s.endpoints.find((e) => e.id === s.activeEndpoint) ?? s.endpoints[0];
}

/**
 * Fælles validering af indstillinger mod en politik (bruges af bridge og direkte tilstand).
 * Returnerer en fejlbesked eller undefined.
 */
export function checkSettings(next: Settings, policy: Policy, isAllowed: (url: string, p: Policy) => boolean): string | undefined {
  if (!next.endpoints.some((e) => e.id === next.activeEndpoint)) return "Aktiv endpoint findes ikke.";
  const ids = new Set<string>();
  for (const e of next.endpoints) {
    if (ids.has(e.id)) return `Endpoint-id '${e.id}' er brugt to gange.`;
    ids.add(e.id);
    if (!isAllowed(e.baseUrl, policy)) return `Host for '${e.name}' (${e.baseUrl}) er ikke tilladt af politikken. Tilføj hosten under "Tilladte hosts".`;
  }
  return undefined;
}
