/**
 * Web-søgning uden API-nøgle (ADR-0010): skraper DuckDuckGo's HTML-endpoint.
 * Kører i bridgen (Node) – browseren har ingen CORS-adgang til endepunktet.
 * Resultaterne sendes til modellen som tekst (utroværdig input – se ADR-0010).
 */

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
  return decodeEntities(s.replace(/<[^>]*>/g, "")).replace(/\s+/g, " ").trim();
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

/** Søger på DuckDuckGo og returnerer et tekstblokk til modellen. Fejl giver en besked, ikke en exception. */
export async function webSearch(query: string, signal?: AbortSignal): Promise<string> {
  const timeout = AbortSignal.timeout(15_000);
  const combined = signal ? AbortSignal.any([signal, timeout]) : timeout;
  try {
    const res = await fetch(`https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`, {
      headers: { "user-agent": UA, accept: "text/html" },
      signal: combined,
    });
    if (!res.ok) return `Søgetjenesten svarede ${res.status} (ingen resultater). Prøv igen senere.`;
    const html = await res.text();
    const results = parseDdgHtml(html).slice(0, 5);
    if (!results.length) return "Ingen resultater fra søgetjenesten.";
    const block = results.map((r, i) => `${i + 1}. ${r.title}\n   ${r.url}\n   ${r.snippet}`).join("\n");
    return `Søgeresultater fra webbet (kan være upålidelige – verificer vigtige fakta):\n${block}`.slice(0, 6_000);
  } catch (err) {
    if ((signal?.aborted ?? false)) throw err; // brugeren afbrød – lad motoren håndtere det
    return `Søgetjenesten svarede ikke (${(err as Error).message}). Prøv igen senere.`;
  }
}
