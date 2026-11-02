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
    instance.initialize({ startOnLoad: false, securityLevel: "strict", theme: "default" });
  }
  return instance;
}

/** Renderer Mermaid-kode som SVG-streng. Kaster, hvis diagrammet er ugyldigt. */
export async function mermaidToSvg(code: string): Promise<string> {
  const m = await ensureMermaid();
  const id = `miai-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  const { svg } = await m.render(id, code);
  return svg;
}
