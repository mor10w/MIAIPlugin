import type { Settings } from "./types.js";

/**
 * Model pr. opgavetype (#13, opfølgning): i stedet for at vælge en model til hver af de mange
 * opgaver vælger brugeren én pr. type. Typerne følger, hvad opgaverne kræver af modellen: korte
 * rettelser klarer en lille, hurtig model; agentens planer har brug for den største.
 * Eval-rapporten anbefaler en model pr. type, så anbefalingen kan sættes direkte ind.
 */
export const MODEL_GROUPS = [
  {
    id: "edit",
    label: { da: "Ret og omskriv", en: "Edit and rewrite" },
    hint: { da: "Stavning, omskrivning, oversættelse, resumé og egne opgaver", en: "Spelling, rewriting, translation, summaries and your own tasks" },
    tasks: ["grammar", "rewrite", "paraphrase", "simplify", "translate", "summarize"],
  },
  {
    id: "write",
    label: { da: "Skriv nyt", en: "Write new" },
    hint: { da: "Artikler, salgstekster, dispositioner og svar på mails", en: "Articles, sales copy, outlines and email replies" },
    tasks: ["expand-article", "sales-copy", "outline", "reply-mail"],
  },
  {
    id: "structure",
    label: { da: "Struktur, formler og data", en: "Structure, formulas and data" },
    hint: { da: "Slides, diagrammer, design, formler og dataanalyse", en: "Slides, diagrams, design, formulas and data analysis" },
    tasks: ["text-to-slides", "diagram", "design-deck", "analyze-data", "explain-formula", "generate-formula"],
  },
  {
    id: "agent",
    label: { da: "Agent", en: "Agent" },
    hint: { da: "Planer, der ændrer dokumentet – også fra chatten", en: "Plans that change the document – also from the chat" },
    tasks: ["agent"],
  },
  {
    id: "chat",
    label: { da: "Chat", en: "Chat" },
    hint: { da: "Spørgsmål og svar om dokumentet", en: "Questions and answers about the document" },
    tasks: ["chat"],
  },
] as const;

export type ModelGroupId = (typeof MODEL_GROUPS)[number]["id"];

/** Typen for en opgave. Egne opgaver (og ukendte) arbejder på markeringen og hører under "Ret og omskriv". */
export function modelGroupOf(taskId: string): ModelGroupId {
  return MODEL_GROUPS.find((g) => (g.tasks as readonly string[]).includes(taskId))?.id ?? "edit";
}

/**
 * Den model, brugeren har valgt til en opgave: et valg for netop opgaven, ellers for dens type,
 * ellers standardmodellen. Tom = intet valgt (motoren tager så den første fra endpointen).
 */
export function chosenModel(settings: Pick<Settings, "taskModels" | "groupModels" | "defaultModel">, taskId: string): string {
  return settings.taskModels[taskId] || settings.groupModels[modelGroupOf(taskId)] || settings.defaultModel || "";
}
