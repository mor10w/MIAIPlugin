import type { ApplyMode, DocContext, GrammarResult } from "@miai/shared";
import { mdToSafeHtml } from "../markdown";
import { truncate, type ApplyOutcome, type ContextOptions, type HostAdapter, type TextFormat } from "./types";

const has = (v: string) => Office.context.requirements.isSetSupported("WordApi", v);

/** Escape af Words specialtegn i søgning (^ bruges til fx ^p). */
const escapeSearch = (s: string) => s.replace(/\^/g, "^^");

export class WordAdapter implements HostAdapter {
  readonly host = "word" as const;
  /** Markeringen fra seneste getContext, sporet så den kan bruges senere selvom brugeren klikker andetsteds. */
  private range: Word.Range | undefined;

  supports(mode: ApplyMode): boolean {
    if (mode === "trackChanges" || mode === "comment") return has("1.4");
    return mode === "replace" || mode === "insertAfter";
  }

  async getContext(opts: ContextOptions): Promise<DocContext> {
    if (this.range) {
      const old = this.range;
      this.range = undefined;
      await Word.run(old, async (ctx) => {
        ctx.trackedObjects.remove(old);
        await ctx.sync();
      }).catch(() => {});
    }
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();
      let target: Word.Range = sel;
      let text = sel.text ?? "";
      let isWholeDocument = false;
      if (!text.trim() && opts.wholeDocumentFallback) {
        const body = ctx.document.body.getRange("Whole");
        body.load("text");
        await ctx.sync();
        target = body;
        text = body.text;
        isWholeDocument = true;
      }
      ctx.trackedObjects.add(target);
      await ctx.sync();
      this.range = target;
      return { host: "word", text: truncate(text), isWholeDocument };
    });
  }

  private run<T>(fn: (ctx: Word.RequestContext, range: Word.Range) => Promise<T>): Promise<T> {
    const r = this.range;
    if (!r) return Word.run((ctx) => fn(ctx, ctx.document.getSelection()));
    return Word.run(r, (ctx) => fn(ctx, r));
  }

  /** Kører fn med sporede ændringer slået til og gendanner derefter brugerens indstilling. */
  private async withTracking<T>(ctx: Word.RequestContext, on: boolean, fn: () => Promise<T>): Promise<T> {
    if (!on) return fn();
    ctx.document.load("changeTrackingMode");
    await ctx.sync();
    const prev = ctx.document.changeTrackingMode;
    ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    try {
      return await fn();
    } finally {
      ctx.document.changeTrackingMode = prev;
      await ctx.sync();
    }
  }

  async applyText(text: string, format: TextFormat, mode: ApplyMode, dc: DocContext): Promise<ApplyOutcome> {
    return this.run(async (ctx, range) => {
      if (mode === "comment") {
        range.insertComment(text);
        await ctx.sync();
        return {};
      }
      const target = mode === "insertAfter" && dc.isWholeDocument ? ctx.document.body.getRange("End") : range;
      const loc = mode === "insertAfter" ? "After" : "Replace";
      await this.withTracking(ctx, mode === "trackChanges", async () => {
        if (format === "markdown") target.insertHtml(mdToSafeHtml(text), loc as "After" | "Replace");
        else target.insertText(mode === "insertAfter" ? `\n${text}` : text, loc as "After" | "Replace");
        await ctx.sync();
      });
      return {};
    });
  }

  async applyGrammar(result: GrammarResult, mode: ApplyMode): Promise<ApplyOutcome> {
    const edits = result.edits.filter((e) => e.original && e.original !== e.replacement && e.original.length <= 255);
    if (!edits.length) return { message: "Ingen rettelser at anvende." };
    return this.run(async (ctx, range) => {
      const searches = edits.map((e) => {
        const res = range.search(escapeSearch(e.original), { matchCase: true });
        res.load("items");
        return res;
      });
      await ctx.sync();
      // Samme original kan forekomme flere gange: brug n'te forekomst for n'te rettelse.
      const seen = new Map<string, number>();
      let applied = 0;
      await this.withTracking(ctx, mode === "trackChanges" && has("1.4"), async () => {
        edits.forEach((e, i) => {
          const n = seen.get(e.original) ?? 0;
          seen.set(e.original, n + 1);
          const hit = searches[i].items[n] ?? searches[i].items[0];
          if (hit) {
            hit.insertText(e.replacement, "Replace");
            applied++;
          }
        });
        await ctx.sync();
      });
      const skipped = edits.length - applied;
      return { message: `${applied} af ${edits.length} rettelser anvendt${mode === "trackChanges" ? " som sporede ændringer" : ""}.`, warning: skipped > 0 };
    });
  }
}
