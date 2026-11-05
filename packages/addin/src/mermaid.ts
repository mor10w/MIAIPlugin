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
    // wrappingWidth: standardbredden (200) deler lange danske ord ("Domænearkite-kt"); 280 holder dem hele.
    instance.initialize({
      startOnLoad: false,
      securityLevel: "strict",
      theme: "default",
      htmlLabels: false,
      flowchart: { htmlLabels: false, wrappingWidth: 280 },
    });
  }
  return instance;
}

/** Cache pr. kode: chatten genrender ved hver ændring, og indsættelse som billede bruger samme SVG. */
const cache = new Map<string, Promise<string>>();
const MAX_CACHE = 50;
/** mermaid.render er ikke sikker ved samtidige kald – kør dem i kø. */
let queue: Promise<unknown> = Promise.resolve();

/** Åbnings- og lukketegn for flowchart-figurer, længste først ("((" før "("). */
const SHAPES: [string, string][] = [
  ["(((", ")))"],
  ["((", "))"],
  ["([", "])"],
  ["[[", "]]"],
  ["[(", ")]"],
  ["{{", "}}"],
  ["[", "]"],
  ["(", ")"],
  ["{", "}"],
];
const INLINE_HTML = /<\/?(?:b|strong|i|em)\s*>/i;

/** Etikettens indhold som mermaid-streng: "…", eller "`…`" (markdown) når den har fed/kursiv. */
function quoteLabel(raw: string): string {
  let text = raw.trim();
  const quoted = /^"([\s\S]*)"$/.exec(text);
  if (quoted) text = quoted[1];
  if (/^`[\s\S]*`$/.test(text)) return `"${text}"`; // allerede en markdown-streng
  if (INLINE_HTML.test(text)) {
    // Uden htmlLabels vises <b> og <i> som tekst – markdown-strenge giver rigtig fed og kursiv.
    const md = text
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/?(?:b|strong)\s*>/gi, "**")
      .replace(/<\/?(?:i|em)\s*>/gi, "*")
      .replace(/<[^>]+>/g, "")
      .replace(/[`"]/g, "'");
    return `"\`${md}\`"`;
  }
  return `"${text.replace(/"/g, "#quot;")}"`;
}

/**
 * Retter de fejl, små modeller oftest laver i flowcharts: etiketter med parenteser, klammer eller
 * andre specialtegn uden anførselstegn (A[Navn (note)] giver en syntaksfejl) og HTML-fremhævning
 * (<b>, <i>), som ikke vises uden htmlLabels. Andre diagramtyper returneres uændret.
 */
export function repairMermaid(code: string): string {
  const lines = code.split("\n");
  if (!/^\s*(?:graph|flowchart)\b/i.test(lines.find((l) => l.trim() && !l.trim().startsWith("%%")) ?? "")) return code;
  return lines
    .map((line) => {
      if (/^\s*(?:%%|classDef|class |style |linkStyle|click |direction )/.test(line)) return line;
      let out = "";
      let i = 0;
      while (i < line.length) {
        // En node-id (eller "subgraph id") efterfulgt af en figur.
        const m = /^(subgraph\s+[\w\-æøåÆØÅ]+\s*|[\w\-æøåÆØÅ]+)/.exec(line.slice(i));
        const prev = i === 0 ? " " : line[i - 1];
        if (m && /[\s;&>|-]/.test(prev)) {
          const after = i + m[0].length;
          const shape = line[after - 1] === "@" ? undefined : SHAPES.find(([open]) => line.startsWith(open, after));
          if (shape) {
            const [open, close] = shape;
            const start = after + open.length;
            // Find lukketegnet på samme niveau: indre (…)/[…]/{…} i teksten tælles med.
            let depth = 0;
            let end = -1;
            let inQuote = false;
            for (let k = start; k < line.length; k++) {
              const ch = line[k];
              if (ch === '"') inQuote = !inQuote;
              if (inQuote) continue;
              if (depth === 0 && line.startsWith(close, k)) {
                end = k;
                break;
              }
              if ("([{".includes(ch)) depth++;
              else if (")]}".includes(ch)) depth = Math.max(0, depth - 1);
            }
            if (end > start) {
              out += m[0] + open + quoteLabel(line.slice(start, end)) + close;
              i = end + close.length;
              continue;
            }
          }
          out += m[0];
          i = after;
          continue;
        }
        out += line[i++];
      }
      return out;
    })
    .join("\n");
}

async function renderRaw(code: string): Promise<string> {
  const m = await ensureMermaid();
  const id = `miai-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  return (await m.render(id, code)).svg;
}

/**
 * Renderer Mermaid-kode som SVG-streng. Kan diagrammet ikke læses, eller indeholder det HTML-
 * fremhævning, prøves den reparerede udgave (repairMermaid). Kaster med mermaids fejl, hvis
 * ingen af dem kan tegnes.
 */
export function mermaidToSvg(code: string): Promise<string> {
  const hit = cache.get(code);
  if (hit) return hit;
  const job = queue.then(async () => {
    const repaired = repairMermaid(code);
    if (repaired === code) return renderRaw(code);
    try {
      return await renderRaw(repaired);
    } catch {
      return renderRaw(code); // reparationen hjalp ikke: vis originalens fejl
    }
  });
  queue = job.catch(() => {});
  // Fejl caches ikke, så et rettet diagram kan prøves igen.
  job.catch(() => cache.delete(code));
  cache.set(code, job);
  if (cache.size > MAX_CACHE) cache.delete(cache.keys().next().value!);
  return job;
}
