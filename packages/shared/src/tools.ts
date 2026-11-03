import type { StructuredOutputMode } from "./types.js";

/**
 * Kendte OpenAI-kompatible inference-værktøjer (#4): standardport, anbefalet indstilling for
 * struktureret output og præcis CORS-opsætning til direkte tilstand.
 */
export interface KnownTool {
  id: string;
  name: string;
  baseUrl: string;
  structuredOutput: StructuredOutputMode;
}

export const KNOWN_TOOLS: KnownTool[] = [
  { id: "ollama", name: "Ollama", baseUrl: "http://127.0.0.1:11434/v1", structuredOutput: "json_schema" },
  { id: "lmstudio", name: "LM Studio", baseUrl: "http://127.0.0.1:1234/v1", structuredOutput: "json_schema" },
  { id: "llamacpp", name: "llama.cpp server", baseUrl: "http://127.0.0.1:8080/v1", structuredOutput: "json_schema" },
  { id: "vllm", name: "vLLM", baseUrl: "http://127.0.0.1:8000/v1", structuredOutput: "json_schema" },
  { id: "localai", name: "LocalAI", baseUrl: "http://127.0.0.1:8080/v1", structuredOutput: "json_object" },
  { id: "jan", name: "Jan", baseUrl: "http://127.0.0.1:1337/v1", structuredOutput: "json_object" },
  { id: "kobold", name: "KoboldCpp", baseUrl: "http://127.0.0.1:5001/v1", structuredOutput: "none" },
];

/** Værktøjer grupperet pr. adresse (llama.cpp og LocalAI deler fx port 8080), til automatisk søgning. */
export function discoveryTargets(): { baseUrl: string; tools: KnownTool[] }[] {
  const byUrl = new Map<string, KnownTool[]>();
  for (const t of KNOWN_TOOLS) byUrl.set(t.baseUrl, [...(byUrl.get(t.baseUrl) ?? []), t]);
  return [...byUrl.entries()].map(([baseUrl, tools]) => ({ baseUrl, tools }));
}

export interface CorsStep {
  /** Fx "Windows", "macOS" eller tom, når trinnet gælder alle. */
  platform?: string;
  /** Kommando eller indstilling, der kan kopieres. */
  code?: string;
  text: string;
}

/**
 * Præcis CORS-opsætning for et værktøj, så add-in'et (på `origin`) må kalde det direkte.
 * Brug altid den præcise origin – ikke "*", som ville lade enhver webside kalde den lokale model.
 */
export function corsSteps(toolId: string, origin: string, lang: "da" | "en"): CorsStep[] {
  const da = lang === "da";
  const restart = (name: string) => (da ? `Genstart ${name} bagefter.` : `Restart ${name} afterwards.`);
  switch (toolId) {
    case "ollama":
      return [
        {
          platform: "Windows",
          code: `setx OLLAMA_ORIGINS "${origin}"`,
          text: `${da ? "Kør i en kommandoprompt." : "Run in a command prompt."} ${restart("Ollama")}`,
        },
        { platform: "macOS", code: `launchctl setenv OLLAMA_ORIGINS "${origin}"`, text: `${da ? "Kør i Terminal." : "Run in Terminal."} ${restart("Ollama")}` },
        {
          platform: "Linux",
          code: `Environment="OLLAMA_ORIGINS=${origin}"`,
          text: da
            ? "Tilføj under [Service] med systemctl edit ollama, og genstart tjenesten."
            : "Add under [Service] with systemctl edit ollama and restart the service.",
        },
      ];
    case "lmstudio":
      return [
        {
          text: da
            ? 'Developer → Server Settings → slå "Enable CORS" til, og genstart serveren.'
            : 'Developer → Server Settings → turn on "Enable CORS" and restart the server.',
        },
      ];
    case "vllm":
      return [{ code: `--allowed-origins '["${origin}"]'`, text: da ? "Tilføj til vllm serve-kommandoen." : "Add to the vllm serve command." }];
    case "localai":
      return [
        {
          code: `CORS=true CORS_ALLOW_ORIGINS=${origin}`,
          text: da ? "Sæt miljøvariablerne, og genstart LocalAI." : "Set the environment variables and restart LocalAI.",
        },
      ];
    case "jan":
      return [{ text: da ? "Settings → Local API Server → slå CORS til." : "Settings → Local API Server → turn on CORS." }];
    case "llamacpp":
    case "kobold":
      return [
        {
          text: da
            ? "Tillader normalt alle origins. Tjek, at serveren lytter på 127.0.0.1 og den viste port."
            : "Usually allows all origins. Check that the server listens on 127.0.0.1 and the port shown.",
        },
      ];
    default:
      return [
        {
          text: da
            ? `Tillad CORS fra ${origin} i værktøjets indstillinger, eller brug MIAI Bridge.`
            : `Allow CORS from ${origin} in the tool's settings, or use MIAI Bridge.`,
        },
      ];
  }
}
