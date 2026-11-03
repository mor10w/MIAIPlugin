import type { AgentAction, ApplyMode, DiagramResult, DocContext, GrammarResult, WordStyle } from "@miai/shared";
import { diagramToPng, graphicToPng, pointSize, type PngImage } from "../image";
import { mdToSafeHtml } from "../markdown";
import { extractGraphic } from "../svg";
import {
  truncate,
  type AgentOptions,
  type AgentOutcome,
  type AgentStepResult,
  type ApplyOutcome,
  type ContextOptions,
  type HostAdapter,
  type TextFormat,
} from "./types";

const has = (v: string) => Office.context.requirements.isSetSupported("WordApi", v);

/** Escape af Words specialtegn i søgning (^ bruges til fx ^p). */
const escapeSearch = (s: string) => s.replace(/\^/g, "^^");

/** Agentens typografier → Words indbyggede (kun indbyggede typografier; ingen fri formatering). */
const STYLE: Record<WordStyle, string> = {
  normal: "Normal",
  heading1: "Heading1",
  heading2: "Heading2",
  heading3: "Heading3",
  bullet: "ListBullet",
};

export class WordAdapter implements HostAdapter {
  readonly host = "word" as const;
  /** Markeringen fra seneste getContext, sporet så den kan bruges senere selvom brugeren klikker andetsteds. */
  private range: Word.Range | undefined;

  supports(mode: ApplyMode): boolean {
    if (mode === "trackChanges" || mode === "comment") return has("1.4");
    return mode === "replace" || mode === "insertAfter" || mode === "diagram" || mode === "image" || mode === "agent";
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
      return { host: "word", ...truncate(text, opts.maxChars), isWholeDocument };
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

  /**
   * Indsætter diagrammet som billede i et nyt afsnit efter markeringen (eller sidst i
   * dokumentet, hvis hele dokumentet blev brugt) – ADR-0009.
   */
  async applyDiagram(result: DiagramResult, dc: DocContext): Promise<ApplyOutcome> {
    const png = await diagramToPng(result);
    return this.insertPicture(png, dc, result.title ?? "Diagram", result.nodes.map((n) => n.label).join(" → "));
  }

  /** Indsætter den første ```mermaid- eller ```svg-blok i teksten som billede efter markeringen. */
  async applyImage(text: string, dc?: DocContext): Promise<ApplyOutcome> {
    const block = extractGraphic(text);
    if (!block) throw new Error("Svaret indeholder ikke et diagram (mermaid eller svg).");
    const png = await graphicToPng(block);
    return this.insertPicture(png, dc, "Figur", block.kind === "mermaid" ? block.code.slice(0, 500) : "");
  }

  /**
   * Fælles indsættelse af billeder: i et nyt afsnit efter markeringen (eller sidst i dokumentet, hvis
   * hele dokumentet blev brugt), i naturlig størrelse – højst A4's tekstbredde – og med alt-tekst.
   */
  private insertPicture(png: PngImage, dc: DocContext | undefined, title: string, description: string): Promise<ApplyOutcome> {
    const { width, height } = pointSize(png);
    return this.run(async (ctx, range) => {
      const para = dc?.isWholeDocument ? ctx.document.body.insertParagraph("", "End") : range.insertParagraph("", "After");
      const pic = para.insertInlinePictureFromBase64(png.base64, "Start");
      pic.width = width;
      pic.height = height;
      pic.altTextTitle = title;
      if (description) pic.altTextDescription = description;
      await ctx.sync();
      return { message: "Figuren er indsat efter markeringen." };
    });
  }
  agentTracking(): boolean {
    return has("1.4");
  }

  /**
   * Agent-tilstand (#6): udfører de bekræftede handlinger én ad gangen inden for markeringen
   * (eller hele dokumentet). Fortryd gendanner dokumentets indhold fra et OOXML-øjebliksbillede
   * taget lige før – også når ændringerne er lavet som sporede ændringer.
   */
  async applyAgent(actions: AgentAction[], dc: DocContext, opts: AgentOptions = {}): Promise<AgentOutcome> {
    return this.run(async (ctx, range) => {
      const body = ctx.document.body;
      const snapshot = body.getOoxml();
      await ctx.sync();
      const ooxml = snapshot.value;
      const scope = dc.isWholeDocument ? body.getRange("Whole") : range;
      const results: AgentStepResult[] = [];
      await this.withTracking(ctx, !!opts.trackChanges && has("1.4"), async () => {
        for (const a of actions) {
          try {
            results.push(await this.agentStep(ctx, scope, dc, a));
          } catch (err) {
            results.push({ ok: false, error: "failed", detail: (err as Error).message });
          }
        }
      });
      const undo = () =>
        Word.run(async (c) => {
          c.document.body.insertOoxml(ooxml, "Replace");
          await c.sync();
        });
      return { results, undo };
    });
  }

  private async find(ctx: Word.RequestContext, scope: Word.Range, text: string): Promise<Word.Range[]> {
    const res = scope.search(escapeSearch(text), { matchCase: true });
    res.load("items");
    await ctx.sync();
    return res.items;
  }

  private async agentStep(ctx: Word.RequestContext, scope: Word.Range, dc: DocContext, a: AgentAction): Promise<AgentStepResult> {
    switch (a.action) {
      case "replace": {
        const hits = await this.find(ctx, scope, a.find);
        if (!hits.length) return { ok: false, error: "not-found" };
        const targets = a.all ? hits : hits.slice(0, 1);
        for (const h of targets) h.insertText(a.replace, "Replace");
        await ctx.sync();
        return { ok: true, count: targets.length };
      }
      case "delete": {
        const [hit] = await this.find(ctx, scope, a.find);
        if (!hit) return { ok: false, error: "not-found" };
        hit.delete();
        await ctx.sync();
        return { ok: true, count: 1 };
      }
      case "comment": {
        const [hit] = await this.find(ctx, scope, a.find);
        if (!hit) return { ok: false, error: "not-found" };
        hit.insertComment(a.text);
        await ctx.sync();
        return { ok: true, count: 1 };
      }
      case "insert": {
        let para: Word.Paragraph;
        if (a.position === "after") {
          const [anchor] = await this.find(ctx, scope, a.anchor ?? "");
          if (!anchor) return { ok: false, error: "anchor-missing" };
          para = has("1.3") ? anchor.paragraphs.getLast().insertParagraph(a.text, "After") : anchor.insertParagraph(a.text, "After");
        } else if (dc.isWholeDocument) {
          para = ctx.document.body.insertParagraph(a.text, a.position === "start" ? "Start" : "End");
        } else {
          para = scope.insertParagraph(a.text, a.position === "start" ? "Before" : "After");
        }
        if (a.style && has("1.3")) para.styleBuiltIn = STYLE[a.style] as Word.BuiltInStyleName;
        await ctx.sync();
        return { ok: true, count: 1 };
      }
      case "move": {
        const [hit] = await this.find(ctx, scope, a.find);
        if (!hit) return { ok: false, error: "not-found" };
        const [anchor] = await this.find(ctx, scope, a.anchor);
        if (!anchor) return { ok: false, error: "anchor-missing" };
        // Flyttes som OOXML, så typografi og tabeller bevares; originalen slettes først,
        // når indsættelsen er lykket (ellers fejler handlingen uden at røre dokumentet).
        const xml = hit.getOoxml();
        await ctx.sync();
        anchor.insertOoxml(xml.value, a.position === "before" ? "Before" : "After");
        hit.delete();
        await ctx.sync();
        return { ok: true, count: 1 };
      }
      default:
        return { ok: false, error: "failed", detail: `${a.action} findes ikke i Word` };
    }
  }
}
