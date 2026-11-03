/**
 * Mermaid-støtte (ADR-0009): modellen skriver ofte ```mermaid-blokke i chatten.
 * De renderes visuelt i task pane og kan indsættes som billede i Word/PowerPoint.
 * Mermaid lazy-importeres, så den ~1 MB store pakke ikke bremser den første rendering.
 */

type MermaidInstance = (typeof import("mermaid"))["default"];

/** Returnerer koden fra den første ```mermaid-blok i teksten (uden gærde), ellers null. */
export function extractMermaid(text: string): string | null {
  const m = /```mermaid[ \t]*\r?\n([\s\S]*?)```/i.exec(text);
  return m ? m[1].trim() : null;
}

let instance: MermaidInstance | undefined;

async function ensureMermaid(): Promise<MermaidInstance> {
  if (!instance) {
    const mod = await import("mermaid");
    instance = mod.default;
    // htmlLabels: false → labels som ren SVG-tekst i stedet for <foreignObject>. WebKit (Office på Mac)
    // "tainter" et canvas, der tegner SVG med foreignObject, så PNG-eksporten til Word/PowerPoint fejler.
    instance.initialize({ startOnLoad: false, securityLevel: "strict", theme: "default", htmlLabels: false, flowchart: { htmlLabels: false } });
  }
  return instance;
}

/** Cache pr. kode: chatten genrender ved hver ændring, og indsættelse som billede bruger samme SVG. */
const cache = new Map<string, Promise<string>>();
const MAX_CACHE = 50;
/** mermaid.render er ikke sikker ved samtidige kald – kør dem i kø. */
let queue: Promise<unknown> = Promise.resolve();

/** Renderer Mermaid-kode som SVG-streng. Kaster, hvis diagrammet er ugyldigt. */
export function mermaidToSvg(code: string): Promise<string> {
  const hit = cache.get(code);
  if (hit) return hit;
  const job = queue.then(async () => {
    const m = await ensureMermaid();
    const id = `miai-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
    return (await m.render(id, code)).svg;
  });
  queue = job.catch(() => {});
  // Fejl caches ikke, så et rettet diagram kan prøves igen.
  job.catch(() => cache.delete(code));
  cache.set(code, job);
  if (cache.size > MAX_CACHE) cache.delete(cache.keys().next().value!);
  return job;
}
