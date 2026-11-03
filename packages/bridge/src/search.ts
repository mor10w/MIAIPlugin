/**
 * Web-søgning uden API-nøgle (ADR-0010): skraper DuckDuckGo's HTML-endpoint.
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
    return await searchDuckDuckGo(query, signal);
  } catch (err) {
    if (signal?.aborted ?? false) throw err; // brugeren afbrød – lad motoren håndtere det
    const host = cfg.provider === "searxng" && cfg.url ? new URL(cfg.url).host : DDG_HOST;
    return { host, ok: false, text: `Søgetjenesten svarede ikke (${(err as Error).message}). Prøv igen senere.` };
  }
}

/** @deprecated Brug searchWeb. Bevares til eksisterende tests (DuckDuckGo). */
export async function webSearch(query: string, signal?: AbortSignal): Promise<string> {
  return (await searchWeb({ provider: "duckduckgo" }, query, () => true, signal)).text;
}
