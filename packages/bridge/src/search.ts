/**
 * Web-søgning uden API-nøgle (ADR-0010): skraper DuckDuckGo's HTML-endpoint, eller bruger en
 * selvhostet SearXNG eller Exas hostede MCP-server.
 * Kører i bridgen (Node) – browseren har ingen CORS-adgang til endepunktet.
 * Resultaterne sendes til modellen som tekst (utroværdig input – se ADR-0010).
 */

import type { SearchConfig } from "@miai/shared";

const UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";

export interface SearchResult {
  title: string;
  url: string;
  snippet: string;
}

function decodeEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#x27;|&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function stripTags(s: string): string {
  return decodeEntities(s.replace(/<[^>]*>/g, ""))
    .replace(/\s+/g, " ")
    .trim();
}

/** DDG linker via //duckduckgo.com/l/?uddg=<url> – udtræk den reelle URL. */
function realUrl(href: string): string {
  const m = /uddg=([^&]+)/.exec(href);
  if (m) {
    try {
      return decodeURIComponent(m[1]);
    } catch {
      // Ugyldig encoding: brug href som den er.
    }
  }
  return href;
}

/** Parser DuckDuckGo's HTML-resultatside. Eksponeret til tests. */
export function parseDdgHtml(html: string): SearchResult[] {
  const results: SearchResult[] = [];
  // Titler og snippets kommer parvis i samme rækkefølge (result__a, result__snippet).
  const titleRe = /<a[^>]*class="result__a"[^>]*>([\s\S]*?)<\/a>/g;
  const snippetRe = /<a[^>]*class="result__snippet"[^>]*>([\s\S]*?)<\/a>/g;
  let m: RegExpExecArray | null;
  while ((m = titleRe.exec(html))) {
    const href = /href="([^"]+)"/.exec(m[0])?.[1] ?? "";
    results.push({ title: stripTags(m[1]), url: realUrl(href), snippet: "" });
  }
  let i = 0;
  while ((m = snippetRe.exec(html)) && i < results.length) {
    results[i++].snippet = stripTags(m[1]);
  }
  return results.filter((r) => r.title && r.url);
}

/** Resultatet af en søgning: tekst til modellen, den faktiske søgehost (til audit-loggen) og om det lykkedes. */
export interface SearchOutcome {
  text: string;
  host: string;
  ok: boolean;
}

export const DDG_HOST = "html.duckduckgo.com";

function formatResults(results: SearchResult[]): string {
  const block = results
    .slice(0, 5)
    .map((r, i) => `${i + 1}. ${r.title}\n   ${r.url}\n   ${r.snippet}`)
    .join("\n");
  return `Søgeresultater fra webbet (kan være upålidelige – verificer vigtige fakta):\n${block}`.slice(0, 6_000);
}

/**
 * Ser siden ud til at indeholde resultater, som parseren ikke kunne læse? Så har DuckDuckGo
 * formentlig ændret sin markup – det skal skelnes fra "ingen resultater" (#12).
 */
export function looksLikeUnparsedResults(html: string): boolean {
  if (/no-results|No results\./i.test(html)) return false;
  return /class="[^"]*\bresult(s|__body|__title)?\b/i.test(html);
}

/**
 * DuckDuckGo's robot-tjek ("bots use DuckDuckGo too"): siden svarer 202 med en opgave i stedet
 * for resultater, typisk efter mange søgninger på kort tid eller fra en VPN-adresse. Det er ikke
 * det samme som "ingen resultater": modellen skal vide, at søgningen ikke blev udført.
 */
export function looksLikeBotChallenge(html: string): boolean {
  return /anomaly-modal|anomaly\.js|bots use DuckDuckGo/i.test(html);
}

/** Beskeden til modellen, når søgetjenesten har afvist søgningen med et robot-tjek. */
export const BOT_CHALLENGE_TEXT =
  "Web-søgningen blev afvist af søgetjenesten (robot-tjek efter for mange søgninger), så der er ingen søgeresultater. Søg ikke igen, og gæt ikke på tal eller fakta: sig, at web-søgningen ikke virkede lige nu.";

function withTimeout(signal?: AbortSignal): AbortSignal {
  const timeout = AbortSignal.timeout(15_000);
  return signal ? AbortSignal.any([signal, timeout]) : timeout;
}

async function searchDuckDuckGo(query: string, signal?: AbortSignal): Promise<SearchOutcome> {
  const host = DDG_HOST;
  const res = await fetch(`https://${DDG_HOST}/html/?q=${encodeURIComponent(query)}`, {
    headers: { "user-agent": UA, accept: "text/html" },
    signal: withTimeout(signal),
  });
  if (!res.ok) return { host, ok: false, text: `Søgetjenesten svarede ${res.status} (ingen resultater). Prøv igen senere.` };
  const html = await res.text();
  const results = parseDdgHtml(html);
  if (results.length) return { host, ok: true, text: formatResults(results) };
  if (res.status === 202 || looksLikeBotChallenge(html)) {
    console.warn("[miai] DuckDuckGo afviste søgningen med et robot-tjek.");
    return { host, ok: false, text: BOT_CHALLENGE_TEXT };
  }
  if (looksLikeUnparsedResults(html)) {
    console.warn("[miai] DuckDuckGo-siden indeholdt resultater, som ikke kunne læses – formatet er måske ændret.");
    return { host, ok: false, text: "Søgetjenestens side kunne ikke læses (formatet er måske ændret). Svar uden søgeresultater, og sig det til brugeren." };
  }
  return { host, ok: true, text: "Ingen resultater fra søgetjenesten." };
}

/** SearXNG's JSON-API (kræver, at "json" er slået til under search.formats i SearXNG's settings.yml). */
async function searchSearxng(baseUrl: string, query: string, signal?: AbortSignal): Promise<SearchOutcome> {
  const host = new URL(baseUrl).host;
  const res = await fetch(`${baseUrl.replace(/\/+$/, "")}/search?q=${encodeURIComponent(query)}&format=json`, {
    headers: { accept: "application/json" },
    signal: withTimeout(signal),
  });
  if (res.status === 403) return { host, ok: false, text: "SearXNG afviste søgningen (403). Er JSON-formatet slået til (search.formats i settings.yml)?" };
  if (!res.ok) return { host, ok: false, text: `Søgetjenesten svarede ${res.status} (ingen resultater). Prøv igen senere.` };
  const json = (await res.json().catch(() => null)) as { results?: { title?: string; url?: string; content?: string }[] } | null;
  if (!json || !Array.isArray(json.results)) return { host, ok: false, text: "Søgetjenesten gav et svar, der ikke kunne læses." };
  const results = json.results
    .filter((r) => r.title && r.url)
    .map((r) => ({ title: r.title!, url: r.url!, snippet: (r.content ?? "").replace(/\s+/g, " ").trim() }));
  return { host, ok: true, text: results.length ? formatResults(results) : "Ingen resultater fra søgetjenesten." };
}

export const EXA_HOST = "mcp.exa.ai";
/** Så meget af Exas svar får modellen: uddragene er lange, og hver opgave kan søge flere gange. */
const EXA_MAX_CHARS = 8_000;
const EXA_RESULTS = 4;

/** Beskeden til modellen, når Exas gratis grænse for søgninger er nået. */
export const EXA_LIMIT_TEXT =
  "Web-søgningen blev afvist af søgetjenesten (grænsen for antal søgninger er nået), så der er ingen søgeresultater. Søg ikke igen, og gæt ikke på tal eller fakta: sig, at web-søgningen ikke virkede lige nu.";

/** Svaret på et JSON-RPC-kald: enten ét JSON-objekt eller en strøm af "data:"-linjer (SSE). */
function parseRpc(body: string): { result?: { content?: { type?: string; text?: string }[]; isError?: boolean }; error?: { message?: string } } | undefined {
  const candidates = body.trimStart().startsWith("{")
    ? [body]
    : body
        .split(/\r?\n/)
        .filter((l) => l.startsWith("data:"))
        .map((l) => l.slice(5).trim());
  // Det sidste objekt med et resultat eller en fejl er svaret; tidligere linjer kan være beskeder undervejs.
  for (const c of candidates.reverse()) {
    try {
      const json = JSON.parse(c) as { result?: unknown; error?: unknown };
      if (json && (json.result || json.error)) return json as ReturnType<typeof parseRpc>;
    } catch {
      // ikke JSON – prøv den forrige linje
    }
  }
  return undefined;
}

/**
 * Exa gennem deres hostede MCP-server (https://mcp.exa.ai/mcp), uden konto og nøgle. Værktøjet
 * web_search_exa kaldes direkte med ét JSON-RPC-kald; serveren kræver ingen session. Svaret er
 * tekst med titel, URL og uddrag af selve siden, så modellen får mere end en kort snippet.
 * Uden nøgle gælder Exas gratis grænse; er den nået, svarer serveren 429.
 */
async function searchExa(query: string, signal?: AbortSignal): Promise<SearchOutcome> {
  const host = EXA_HOST;
  const res = await fetch(`https://${EXA_HOST}/mcp`, {
    method: "POST",
    headers: { "content-type": "application/json", accept: "application/json, text/event-stream" },
    body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/call", params: { name: "web_search_exa", arguments: { query, numResults: EXA_RESULTS } } }),
    signal: withTimeout(signal),
  });
  if (res.status === 429) return { host, ok: false, text: EXA_LIMIT_TEXT };
  if (!res.ok) return { host, ok: false, text: `Søgetjenesten svarede ${res.status} (ingen resultater). Prøv igen senere.` };
  const rpc = parseRpc(await res.text());
  if (!rpc) return { host, ok: false, text: "Søgetjenesten gav et svar, der ikke kunne læses." };
  const text = (rpc.result?.content ?? [])
    .filter((c) => c.type === "text" && c.text)
    .map((c) => c.text!.trim())
    .join("\n\n");
  if (rpc.error || rpc.result?.isError) {
    const why = (rpc.error?.message ?? text).slice(0, 300);
    // Grænsen kan også komme som en fejl i selve svaret.
    if (/rate.?limit|too many requests|\b429\b/i.test(why)) return { host, ok: false, text: EXA_LIMIT_TEXT };
    return { host, ok: false, text: `Søgetjenesten meldte en fejl (${why || "ukendt"}). Svar uden søgeresultater, og sig det til brugeren.` };
  }
  if (!text) return { host, ok: true, text: "Ingen resultater fra søgetjenesten." };
  return { host, ok: true, text: `Søgeresultater fra webbet (kan være upålidelige – verificer vigtige fakta):\n${text}`.slice(0, EXA_MAX_CHARS) };
}

/**
 * Søger med den konfigurerede tjeneste og returnerer en tekstblok til modellen.
 * Fejl giver en besked til modellen, ikke en exception (undtagen når brugeren afbryder).
 * `isAllowed` håndhæver politikken for en selvhostet tjeneste (samme regel som model-endpoints).
 */
export async function searchWeb(cfg: SearchConfig, query: string, isAllowed: (url: string) => boolean, signal?: AbortSignal): Promise<SearchOutcome> {
  try {
    if (cfg.provider === "searxng") {
      if (!cfg.url) return { host: "searxng", ok: false, text: "Ingen SearXNG-adresse er sat op." };
      if (!isAllowed(cfg.url)) return { host: new URL(cfg.url).host, ok: false, text: "Søgetjenesten er ikke tilladt af politikken." };
      return await searchSearxng(cfg.url, query, signal);
    }
    if (cfg.provider === "exa") return await searchExa(query, signal);
    return await searchDuckDuckGo(query, signal);
  } catch (err) {
    if (signal?.aborted ?? false) throw err; // brugeren afbrød – lad motoren håndtere det
    const host = cfg.provider === "searxng" && cfg.url ? new URL(cfg.url).host : cfg.provider === "exa" ? EXA_HOST : DDG_HOST;
    return { host, ok: false, text: `Søgetjenesten svarede ikke (${(err as Error).message}). Prøv igen senere.` };
  }
}

/** @deprecated Brug searchWeb. Bevares til eksisterende tests (DuckDuckGo). */
export async function webSearch(query: string, signal?: AbortSignal): Promise<string> {
  return (await searchWeb({ provider: "duckduckgo" }, query, () => true, signal)).text;
}
