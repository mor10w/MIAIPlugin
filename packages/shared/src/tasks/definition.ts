/** Typerne for en opgave i kataloget (#60): hvad den tager, hvad den giver, og hvordan svaret bruges. */
import type { z } from "zod";
import type { ChatMessage, DocContext, HostKind, ParsedTaskRequest } from "../types.js";

export type ApplyMode =
  | "replace" // erstat markeringen
  | "trackChanges" // erstat som sporede ændringer (Word)
  | "insertAfter" // indsæt efter markeringen
  | "comment" // indsæt som kommentar
  | "image" // indsæt Mermaid-diagram i svaret som billede (Word/PowerPoint)
  | "newSlides" // opret slides (PowerPoint)
  | "formula" // skriv formel i markeret celle (Excel)
  | "cells" // skriv TSV ind i markerede celler (Excel)
  | "diagram" // indsæt diagram (billede i Word, figurer i PowerPoint)
  | "agent" // udfør de valgte handlinger fra agentens plan (#6)
  | "design" // opret præsentationens master og slides ud fra designet (PowerPoint, ADR-0014)
  | "reply" // åbn et svar på mailen med teksten (Outlook, læsetilstand)
  | "replyAll"; // åbn et svar til alle (Outlook, læsetilstand)

export type Localized = { da: string; en: string };

export type TextFormat = "plain" | "markdown" | "tsv";

/** Input er et markeret Excel-område med værdier (bruges som TSV i prompten). */
export function isExcelRange(ctx: DocContext): boolean {
  return ctx.host === "excel" && !!ctx.table;
}

export interface ParamOption {
  value: string;
  label: Localized;
}

export interface ParamDef {
  name: string;
  label: Localized;
  kind: "select" | "text";
  options?: ParamOption[];
  default?: string;
  placeholder?: Localized;
}

export type TaskOutput =
  | {
      kind: "text";
      format: TextFormat;
      /** Svaret er TSV, når input er et markeret Excel-område (fx oversættelse celle for celle). */
      tsvForExcelRanges?: boolean;
    }
  | { kind: "json"; schema: z.ZodType; schemaName: string };

/**
 * - selection: kræver markeret tekst
 * - selectionOrDocument: bruger markeringen, ellers hele dokumentet
 * - prompt: brugerens instruktion i params.prompt; markering er valgfri kontekst
 * - formula: Excel-formel i markeret celle
 */
export type TaskInput = "selection" | "selectionOrDocument" | "prompt" | "formula";

export interface TaskDefinition {
  id: string;
  label: Localized;
  description: Localized;
  hosts: HostKind[];
  input: TaskInput;
  output: TaskOutput;
  params: ParamDef[];
  applyModes: ApplyMode[];
  /**
   * map-reduce: lange tekster deles, hver del behandles for sig, og resultaterne samles i en
   * afsluttende kørsel (#3). Add-in'et sender så mere tekst end normalt.
   */
  chunking?: "map-reduce";
  /** Lav temperatur for rettelser, højere for kreativ tekst. Overstyrer indstillingen, hvis sat. */
  temperature?: number;
  /** Fast seed, så samme forespørgsel giver samme svar, hvor serveren understøtter det (#80). */
  seed?: number;
  /**
   * Justering pr. forespørgsel (#80): en opgave kan fravælge værktøjer og tænkning og sætte et
   * lavere loft over svarlængden, fx når en regnearksfunktion kører over mange celler.
   */
  tuning?(req: ParsedTaskRequest): TaskTuning;
  buildMessages(req: ParsedTaskRequest): ChatMessage[];
}

/** Opgaven bag regnearksfunktionen =MIAI.ASK() (#80). */
export const CELL_ASK_TASK = "cell-ask";
/** Loft over svarlængden, når funktionen kaldes med optimized = SAND. */
export const CELL_ASK_OPTIMIZED_TOKENS = 256;

export interface TaskTuning {
  /** Ingen værktøjer (web-søgning), uanset indstillingen. */
  noTools?: boolean;
  /** Slå tænkning fra, uanset indstillingen. */
  noThinking?: boolean;
  /** Loft over svarlængden i tokens (aldrig over indstillingen). */
  maxTokens?: number;
}
