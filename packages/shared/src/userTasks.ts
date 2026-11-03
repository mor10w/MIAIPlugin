import { z } from "zod";
import type { ChatMessage, DocContext, HostKind, ParsedTaskRequest } from "./types.js";
import { doc, hostName, isExcelRange, SYSTEM_BASE, type ApplyMode, type Localized, type ParamDef, type TaskDefinition, type TextFormat } from "./tasks.js";

/**
 * Brugerdefinerede opgaver (F14): opgaver som brugerredigerbare JSON-filer, der deles via mappe.
 * Filen deklareres (prompt-skabelon med {{placeholderer}}) og omsættes til en TaskDefinition,
 * så den fælles motor kører den på samme måde som de indbyggede opgaver.
 */

/** Lokalisering i brugerrfiler: "en" er valgfri og falder tilbage til "da". */
export const UserLocalized = z.object({ da: z.string().min(1).max(200), en: z.string().max(200).optional() });

export const UserParamDef = z.object({
  name: z.string().regex(/^[a-zA-Z][a-zA-Z0-9_]{0,31}$/),
  label: UserLocalized,
  kind: z.enum(["select", "text"]),
  options: z.array(z.object({ value: z.string().min(1).max(200), label: UserLocalized })).min(1).max(20).optional(),
  default: z.string().max(500).optional(),
});

/** Én brugerdefineret opgave. Filnavnet i mappen er <id>.json. */
export const UserTaskFile = z.object({
  /** Unikt id, lille bogstaver/tal/bindestreger. Må ikke kollidere med en indbygget opgave. */
  id: z.string().regex(/^[a-z0-9][a-z0-9-]{0,39}$/),
  label: UserLocalized,
  description: UserLocalized.optional(),
  hosts: z.array(z.enum(["word", "excel", "powerpoint"])).min(1).max(3),
  /** formula kræver JSON-output og kan derfor ikke bruges i filer. */
  input: z.enum(["selection", "selectionOrDocument", "prompt"]),
  output: z.object({ format: z.enum(["plain", "markdown", "tsv"]) }),
  params: z.array(UserParamDef).max(8).default([]),
  /** "cells" bruges til TSV-output i Excel (celle for celle). */
  applyModes: z.array(z.enum(["replace", "trackChanges", "insertAfter", "comment", "cells"])).min(1).max(5),
  temperature: z.number().min(0).max(2).optional(),
  prompt: z.object({
    /** Valgfri ekstra systemprompt (udover MIAI's basisregler). */
    system: z.string().max(8_000).default(""),
    /** Skabelon for brugerbilledet. Mindst ét {{placeholder}} kræves ikke – tom skabelon = kun kontekst. */
    user: z.string().min(1).max(16_000),
  }),
});
export type UserTaskFile = z.infer<typeof UserTaskFile>;

/** Placeholderer tilgængelige i skabelonen ud over opgavens parametre. */
export const TEMPLATE_PLACEHOLDERS = ["selection", "prompt", "language", "address", "table"] as const;

/** Udfylder {{navn}}-placeholderer. Ukendte placeholderer står ubehandlede, så brugeren opdager dem. */
export function renderTemplate(template: string, values: Record<string, string>): string {
  return template.replace(/\{\{(\w+)\}\}/g, (m, name: string) => (name in values ? values[name] : m));
}

/** "en" mangler → brug "da". */
export function localize(l: { da: string; en?: string }, lang: "da" | "en"): string {
  return (lang === "en" ? l.en : undefined) ?? l.da;
}

const fill = (l: { da: string; en?: string }): Localized => ({ da: l.da, en: l.en ?? l.da });

/** Omsætter en brugerdefineret opgave til den form, motoren kender. */
export function toTaskDefinition(spec: UserTaskFile): TaskDefinition {
  return {
    id: spec.id,
    label: fill(spec.label),
    description: fill(spec.description ?? { da: "" }),
    hosts: [...spec.hosts],
    input: spec.input,
    output: { kind: "text", format: spec.output.format },
    params: spec.params.map((p) => ({
      name: p.name,
      label: fill(p.label),
      kind: p.kind,
      options: p.options?.map((o) => ({ value: o.value, label: fill(o.label) })),
      default: p.default,
    })),
    // TSV-output hører hjemme i Excel-celler; sørg for, at knappen findes.
    applyModes: spec.output.format === "tsv" && !spec.applyModes.includes("cells") ? [...spec.applyModes, "cells"] : [...spec.applyModes],
    temperature: spec.temperature,
    buildMessages(req: ParsedTaskRequest): ChatMessage[] {
      const ctx = req.context;
      // Markeringen pakkes i <document>-tags, når skabelonen bruger {{selection}} –
      // samme prompt-injection-sikring som de indbyggede opgaver.
      const usesSelection = spec.prompt.user.includes("{{selection}}") || spec.prompt.system.includes("{{selection}}");
      const values: Record<string, string> = {
        ...req.params,
        selection: usesSelection ? doc(ctx.text) : ctx.text,
        prompt: req.params.prompt ?? "",
        language: req.params.language ?? "English",
        address: ctx.address ?? "",
        // Tabellen er dokumentindhold og pakkes som markeringen (prompt-injection-sikring).
        table: ctx.table ? doc(ctx.table) : "",
      };
      const sysParts = [SYSTEM_BASE, `The user is working in ${hostName(ctx.host)}.`];
      if (spec.prompt.system.trim()) sysParts.push(renderTemplate(spec.prompt.system, values));
      return [
        { role: "system", content: sysParts.join("\n") },
        { role: "user", content: renderTemplate(spec.prompt.user, values) },
      ];
    },
  };
}

/** Serializable opgavekatalog-view, som add-in'et bruger (i stedet for TaskDefinition-funktioner). */
export interface TaskView {
  id: string;
  label: Localized;
  description: Localized;
  hosts: HostKind[];
  input: TaskDefinition["input"];
  output: { kind: "text"; format: TextFormat; tsvForExcelRanges?: boolean } | { kind: "json" };
  params: ParamDef[];
  applyModes: ApplyMode[];
  /** true = brugerdefineret (fil), false = indbygget. */
  custom: boolean;
  chunking?: "map-reduce";
}

export function toTaskView(t: TaskDefinition, custom = false): TaskView {
  return {
    id: t.id,
    label: t.label,
    description: t.description,
    hosts: [...t.hosts],
    input: t.input,
    output: t.output.kind === "text" ? { kind: "text", format: t.output.format, tsvForExcelRanges: t.output.tsvForExcelRanges } : { kind: "json" },
    params: [...t.params],
    applyModes: [...t.applyModes],
    custom,
    chunking: t.chunking,
  };
}

/** Det faktiske tekstformat for en TaskView i en given kontekst (spejler tasks.ts' textFormat). */
export function formatForView(view: TaskView, ctx: DocContext): TextFormat | undefined {
  if (view.output.kind !== "text") return undefined;
  if (view.output.tsvForExcelRanges && isExcelRange(ctx)) return "tsv";
  return view.output.format;
}
