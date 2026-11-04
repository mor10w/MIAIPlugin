import type { AgentAction, ApplyMode, DesignPlan, DiagramResult, DocContext, FormulaResult, GrammarResult, HostKind, SlidesResult } from "@miai/shared";

export type TextFormat = "plain" | "markdown" | "tsv";

export interface ApplyOutcome {
  /** Kald for at fortryde (bruges hvor Office ikke selv kan fortryde add-in-ændringer, fx Excel). */
  undo?: () => Promise<void>;
  /** Besked til brugeren, fx "3 af 4 rettelser anvendt" eller en formelfejl. */
  message?: string;
  warning?: boolean;
}

/** Udfaldet af én agent-handling (#6). */
export interface AgentStepResult {
  ok: boolean;
  /** Antal steder, handlingen ramte (fx forekomster erstattet eller celler skrevet). */
  count?: number;
  /** Hvorfor handlingen ikke blev udført – oversættes i UI'et. */
  error?: "not-found" | "anchor-missing" | "other-sheet" | "formula-error" | "changed" | "no-structure" | "failed";
  /** Teknisk detalje, fx Office' fejlbesked eller Excels fejlværdi (#VALUE!). */
  detail?: string;
}

export interface AgentOutcome extends ApplyOutcome {
  /** Ét resultat pr. udført handling, i samme rækkefølge. */
  results: AgentStepResult[];
  /** Fortryder hele agent-sessionen. */
  undo: () => Promise<void>;
  /**
   * Hvordan fortryd virker: tracked = afviser sessionens sporede ændringer (resten af dokumentet
   * røres ikke); snapshot = gendanner dokumentets indhold fra før sessionen; cells = Excel-celler.
   */
  undoKind: "tracked" | "snapshot" | "cells" | "slides" | "text" | "script";
  /** Kodetilstand: scriptet fejlede (og er rullet tilbage) – fejlen og koden kan sendes til modellen. */
  retry?: { error: string; code: string };
}

export interface DesignOutcome {
  /** Antal indsatte slides. */
  inserted: number;
  /** Antal fjernede (gamle) slides. */
  removed: number;
  /** Sat, hvis de gamle slides ikke kunne fjernes (fx ældre PowerPoint). */
  keptOld?: boolean;
  undo: () => Promise<void>;
}

export interface AgentOptions {
  /** Word: udfør som sporede ændringer, så de også kan gennemgås i Word. */
  trackChanges?: boolean;
  /** Excel: kodetilstanden er slået til og tilladt af politikken (ADR-0017). */
  scripts?: boolean;
}

export interface ContextOptions {
  /** Brug hele dokumentet, hvis intet er markeret. */
  wholeDocumentFallback: boolean;
  /**
   * Excel: hvor meget af arket modellen skal se ud over markeringen.
   * "headers" (standard) = overskrifter og første rækker – billigt.
   * "stats" = kolonnestatistik over hele used range (op til 5.000 rækker) – kun når opgaven bruger den.
   */
  overview?: "headers" | "stats";
  /** Word: medtag afsnittene med typografi, liste og tabel (agent-tilstand, kræver WordApi 1.3). */
  structure?: boolean;
  /** Maks. antal tegn tekst (standard MAX_CONTEXT_CHARS; opgaver med map-reduce får mere). */
  maxChars?: number;
}

export interface HostAdapter {
  readonly host: HostKind;
  /** Hvilke indsætningsmåder host'en understøtter. */
  supports(mode: ApplyMode): boolean;
  getContext(opts: ContextOptions): Promise<DocContext>;
  applyText(text: string, format: TextFormat, mode: ApplyMode, ctx: DocContext): Promise<ApplyOutcome>;
  applyGrammar?(result: GrammarResult, mode: ApplyMode, ctx: DocContext): Promise<ApplyOutcome>;
  applySlides?(result: SlidesResult): Promise<ApplyOutcome>;
  applyFormula?(result: FormulaResult, ctx: DocContext): Promise<ApplyOutcome>;
  /** Indsætter diagrammet visuelt (billede i Word, native shapes i PowerPoint). */
  applyDiagram?(result: DiagramResult, ctx: DocContext): Promise<ApplyOutcome>;
  /** Indsætter den første ```mermaid- eller ```svg-blok i teksten som billede. Kaster, hvis ingen blok findes. */
  applyImage?(text: string, ctx?: DocContext): Promise<ApplyOutcome>;
  /** Agent-tilstand (#6): udfører de bekræftede handlinger og returnerer fortryd for hele sessionen. */
  applyAgent?(actions: AgentAction[], ctx: DocContext, opts?: AgentOptions): Promise<AgentOutcome>;
  /**
   * PowerPoint (ADR-0014): sætter designets master og slides ind. replace = fjern de nuværende
   * slides bagefter (kræver PowerPointApi 1.8, så det kan fortrydes).
   */
  applyDesign?(plan: DesignPlan, opts: { replace: boolean; lang: "da" | "en" }): Promise<DesignOutcome>;
  /** Kan agentens ændringer udføres som sporede ændringer? */
  agentTracking?(): boolean;
}

/** Ca. 6-8k tokens. Større dokumenter afkortes (chunking kommer i fase 1). */
export const MAX_CONTEXT_CHARS = 24_000;

/** Opgaver, der deler lange tekster op (map-reduce), kan få så meget (≈ 60 sider; DocContext' grænse). */
export const MAX_CHUNKED_CHARS = 200_000;

/** Afkorter teksten og fortæller, om det skete, så UI'et kan vise det (#3). */
export function truncate(text: string, maxChars = MAX_CONTEXT_CHARS): { text: string; truncated?: true; originalChars?: number } {
  if (text.length <= maxChars) return { text };
  return { text: `${text.slice(0, maxChars)}\n[… afkortet]`, truncated: true, originalChars: text.length };
}
