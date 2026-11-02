import type { ApplyMode, DocContext, FormulaResult, GrammarResult, HostKind, SlidesResult } from "@miai/shared";

export type TextFormat = "plain" | "markdown" | "tsv";

export interface ApplyOutcome {
  /** Kald for at fortryde (bruges hvor Office ikke selv kan fortryde add-in-ændringer, fx Excel). */
  undo?: () => Promise<void>;
  /** Besked til brugeren, fx "3 af 4 rettelser anvendt" eller en formelfejl. */
  message?: string;
  warning?: boolean;
}

export interface ContextOptions {
  /** Brug hele dokumentet, hvis intet er markeret. */
  wholeDocumentFallback: boolean;
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
}

/** Ca. 6-8k tokens. Større dokumenter afkortes (chunking kommer i fase 1). */
export const MAX_CONTEXT_CHARS = 24_000;

export function truncate(text: string): string {
  return text.length > MAX_CONTEXT_CHARS ? `${text.slice(0, MAX_CONTEXT_CHARS)}\n[… afkortet]` : text;
}
