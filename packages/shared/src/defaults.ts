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
  // SearXNG-adressen håndhæves som model-endpoints (#12). En låst søgetjeneste i politikken valideres ikke her.
  if (!policy.search && next.search.provider === "searxng") {
    if (!next.search.url) return "Angiv adressen på din SearXNG.";
    if (!isAllowed(next.search.url, policy)) return `Søgetjenesten (${next.search.url}) er ikke tilladt af politikken. Tilføj hosten under "Tilladte hosts".`;
  }
  // Lokal RAG (F15): mapperne tjekkes mod politikken i bridgen, som kender filsystemet.
  if (next.knowledge.enabled && !policy.knowledge) return "Søgning i egne mapper er slået fra af din organisation.";
  return undefined;
}
