import { selectedText, wordCount } from "./selection";
import type { SelectionInfo } from "./types";
import type { DocRef } from "../refs";
import {
  minimalEdit,
  occurrences,
  parseBlocks,
  type AgentAction,
  type ApplyMode,
  type DiagramResult,
  type DocContext,
  type GrammarResult,
  type InlineText,
  type WordStyle,
} from "@miai/shared";
import { diagramToPng, graphicToPng, pointSize, type PngImage } from "../image";
import { mdToSafeHtml } from "../markdown";
import { extractGraphic } from "../svg";
import {
  MAX_CONTEXT_CHARS,
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
  numbered: "ListNumber",
  quote: "Quote",
};

/** Agent-sessionens afsnit: som modellen så dem, ændret før start (stale) og slettet undervejs (gone). */
interface Session {
  paras: Word.Paragraph[];
  stale: Set<number>;
  gone: Set<number>;
}

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
      const paragraphs = opts.structure && has("1.3") ? await this.structure(ctx, target, opts.maxChars) : undefined;
      return { host: "word", ...truncate(text, opts.maxChars), isWholeDocument, ...(paragraphs ? { paragraphs } : {}) };
    });
  }

  /**
   * Afsnittene i området med typografi, liste og tabel – så agenten kan se (og bevare)
   * dokumentets struktur. Afkortes ved samme tegngrænse som teksten.
   */
  private async structure(ctx: Word.RequestContext, target: Word.Range, maxChars = MAX_CONTEXT_CHARS): Promise<NonNullable<DocContext["paragraphs"]>> {
    const paras = target.paragraphs;
    paras.load("items/text,items/style,items/styleBuiltIn,items/isListItem,items/tableNestingLevel");
    await ctx.sync();
    const out: NonNullable<DocContext["paragraphs"]> = [];
    let chars = 0;
    for (const p of paras.items) {
      chars += p.text.length + 1;
      if (chars > maxChars) break;
      out.push({
        text: p.text,
        // Indbygget navn (fx "Heading2") er det samme på alle sprog; ellers dokumentets eget typografinavn.
        style: p.styleBuiltIn && p.styleBuiltIn !== "Other" ? p.styleBuiltIn : p.style || "Normal",
        ...(p.isListItem ? { list: true } : {}),
        ...(p.tableNestingLevel > 0 ? { table: true } : {}),
      });
    }
    return out;
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

  /** Antal markerede ord, eller null når intet er markeret (#26). */
  async selectionInfo(): Promise<SelectionInfo | null> {
    const words = wordCount(await selectedText());
    return words ? { kind: "text", words } : null;
  }

  /** Markerer afsnittet eller teksten i dokumentet (#22). Ændrer intet. */
  async goTo(ref: DocRef): Promise<void> {
    await Word.run(async (ctx) => {
      const scope = ctx.document.body.getRange();
      if (ref.paragraph) {
        const paras = scope.paragraphs;
        paras.load("items");
        await ctx.sync();
        paras.items[ref.paragraph - 1]?.select();
      } else if (ref.find) {
        const hits = await this.find(ctx, scope, ref.find.slice(0, 255));
        hits[0]?.select();
      }
      await ctx.sync();
    });
  }

  /**
   * Agent-tilstand (#6, ADR-0012): udfører de bekræftede handlinger én ad gangen og bevarer
   * dokumentets formatering:
   *  - afsnit adresseres med de numre, modellen så (¶), via afsnitsobjekter hentet før første
   *    ændring – så numrene gælder, selv om tidligere handlinger har indsat eller slettet afsnit;
   *  - erstat ændrer kun de ord, der faktisk er forskellige;
   *  - ny tekst får en eksplicit typografi og arver ikke naboens formatering;
   *  - hele afsnit slettes og flyttes som afsnit (ingen tomme afsnitsmærker).
   *
   * Fortryd afviser sessionens sporede ændringer (WordApi 1.6), så resten af dokumentet ikke
   * røres; uden sporede ændringer gendannes indholdet fra et øjebliksbillede.
   */
  async applyAgent(actions: AgentAction[], dc: DocContext, opts: AgentOptions = {}): Promise<AgentOutcome> {
    return this.run(async (ctx, range) => {
      const body = ctx.document.body;
      const tracking = !!opts.trackChanges && has("1.4");
      const undoKind = tracking && has("1.6") ? "tracked" : "snapshot";
      const snapshot = undoKind === "snapshot" ? body.getOoxml() : undefined;
      const scope = dc.isWholeDocument ? body.getRange("Whole") : range;
      // Afsnittene, som modellen så dem (samme rækkefølge som DocContext.paragraphs).
      const paraList = has("1.3") ? scope.paragraphs : undefined;
      paraList?.load("items/text");
      await ctx.sync();
      const ooxml = snapshot?.value;
      const paras = paraList?.items ?? [];
      // Er dokumentet ændret siden planen? Tjekkes én gang før første handling – bagefter ændrer
      // planens egne handlinger afsnittene, og det er forventet.
      const stale = new Set<number>();
      dc.paragraphs?.forEach((seen, i) => {
        if (paras[i]?.text.trim() !== seen.text.trim()) stale.add(i + 1);
      });
      const gone = new Set<number>();
      const results: AgentStepResult[] = [];
      const commentIds: string[] = [];
      const started = Date.now();
      await this.withTracking(ctx, tracking, async () => {
        for (const a of actions) {
          try {
            results.push(await this.agentStep(ctx, scope, dc, { paras, stale, gone }, a, commentIds));
          } catch (err) {
            results.push({ ok: false, error: "failed", detail: (err as Error).message });
          }
        }
      });
      const ended = Date.now();

      const undo = () =>
        Word.run(async (c) => {
          // Agentens kommentarer fjernes i begge tilfælde (de er ikke sporede ændringer).
          if (commentIds.length && has("1.4")) {
            const comments = c.document.body.getComments();
            comments.load("items/id");
            await c.sync();
            for (const cm of comments.items) if (commentIds.includes(cm.id)) cm.delete();
          }
          if (undoKind === "tracked") {
            // Kun ændringer lavet under sessionen (±2 s) – brugerens egne sporede ændringer røres ikke.
            const changes = c.document.body.getTrackedChanges();
            changes.load("items/date");
            await c.sync();
            for (const ch of changes.items) {
              const t = new Date(ch.date).getTime();
              if (t >= started - 2_000 && t <= ended + 2_000) ch.reject();
            }
          } else if (ooxml) {
            c.document.body.insertOoxml(ooxml, "Replace");
          }
          await c.sync();
        });
      return { results, undo, undoKind };
    });
  }

  private async find(ctx: Word.RequestContext, scope: Word.Range, text: string): Promise<Word.Range[]> {
    const res = scope.search(escapeSearch(text), { matchCase: true });
    res.load("items");
    await ctx.sync();
    return res.items;
  }

  /**
   * Afsnit ¶n, som modellen så det – eller en fejl, hvis dokumentet var ændret, før agenten gik i
   * gang, eller afsnittet er slettet af en tidligere handling i planen.
   */
  private paragraph(dc: DocContext, s: Session, n: number | undefined): Word.Paragraph | AgentStepResult {
    if (!has("1.3") || !dc.paragraphs) return { ok: false, error: "no-structure" };
    const p = n !== undefined ? s.paras[n - 1] : undefined;
    if (!p || n === undefined || s.stale.has(n)) return { ok: false, error: "changed" };
    if (s.gone.has(n)) return { ok: false, error: "failed", detail: `¶${n} er slettet af en tidligere handling` };
    return p;
  }

  /**
   * Giver et nyt afsnit en eksplicit typografi og fjerner formatering, det har arvet fra naboen
   * (Word kopierer naboafsnittets egenskaber ved indsættelse), og sætter derefter fed, kursiv og links.
   */
  private async formatNew(ctx: Word.RequestContext, p: Word.Paragraph, block: { style: WordStyle; inline: InlineText }): Promise<void> {
    if (!has("1.3")) return;
    p.styleBuiltIn = STYLE[block.style] as Word.BuiltInStyleName;
    // Brødtekst og lister: nulstil arvet fed/kursiv/understregning (overskrifter og citater får deres fra typografien).
    if (block.style === "normal" || block.style === "bullet" || block.style === "numbered")
      p.getRange("Whole").font.set({ bold: false, italic: false, underline: "None", strikeThrough: false });
    p.load("isListItem");
    await ctx.sync();
    if (p.isListItem && block.style !== "bullet" && block.style !== "numbered") p.detachFromList();
    const { text, spans } = block.inline;
    for (const span of spans) {
      const seg = text.slice(span.start, span.end);
      if (!seg.trim() || seg.length > 255) continue;
      const hits = p.search(escapeSearch(seg), { matchCase: true });
      hits.load("items");
      await ctx.sync();
      const r = hits.items[occurrences(text.slice(0, span.start), seg)] ?? hits.items[0];
      if (!r) continue;
      if (span.bold) r.font.bold = true;
      if (span.italic) r.font.italic = true;
      if (span.link) r.hyperlink = span.link;
    }
    await ctx.sync();
  }

  private async agentStep(
    ctx: Word.RequestContext,
    scope: Word.Range,
    dc: DocContext,
    session: Session,
    a: AgentAction,
    commentIds: string[],
  ): Promise<AgentStepResult> {
    switch (a.action) {
      case "replace": {
        const hits = await this.find(ctx, scope, a.find);
        if (!hits.length) return { ok: false, error: "not-found" };
        const targets = a.all ? hits : hits.slice(0, 1);
        // Kun den del, der faktisk ændres, erstattes – formatering før og efter bevares.
        const { before, oldCore, newCore } = minimalEdit(a.find, a.replace);
        if (!oldCore) return { ok: true, count: 0 };
        for (const h of targets) {
          let target: Word.Range = h;
          if (oldCore !== a.find && oldCore.length <= 255) {
            const sub = h.search(escapeSearch(oldCore), { matchCase: true });
            sub.load("items");
            await ctx.sync();
            target = sub.items[occurrences(before, oldCore)] ?? sub.items[0] ?? h;
          }
          const replacement = target === h ? a.replace : newCore;
          if (replacement) target.insertText(replacement, "Replace");
          else target.delete();
        }
        await ctx.sync();
        return { ok: true, count: targets.length };
      }
      case "delete": {
        if (a.paragraph !== undefined) {
          const p = this.paragraph(dc, session, a.paragraph);
          if (!("delete" in p)) return p;
          p.delete();
          session.gone.add(a.paragraph);
          await ctx.sync();
          return { ok: true, count: 1 };
        }
        const [hit] = await this.find(ctx, scope, a.find ?? "");
        if (!hit) return { ok: false, error: "not-found" };
        // Er det hele afsnittets tekst, slettes afsnittet – ellers bliver et tomt afsnit (eller punkt) stående.
        if (has("1.3")) {
          const para = hit.paragraphs.getFirst();
          para.load("text");
          await ctx.sync();
          if (para.text.trim() === (a.find ?? "").trim()) {
            para.delete();
            await ctx.sync();
            return { ok: true, count: 1 };
          }
        }
        hit.delete();
        await ctx.sync();
        return { ok: true, count: 1 };
      }
      case "comment": {
        const [hit] = await this.find(ctx, scope, a.find);
        if (!hit) return { ok: false, error: "not-found" };
        const c = hit.insertComment(a.text);
        c.load("id");
        await ctx.sync();
        commentIds.push(c.id);
        return { ok: true, count: 1 };
      }
      case "setStyle": {
        const p = this.paragraph(dc, session, a.paragraph);
        if (!("delete" in p)) return p;
        p.styleBuiltIn = STYLE[a.style] as Word.BuiltInStyleName;
        p.load("isListItem");
        await ctx.sync();
        if (p.isListItem && a.style !== "bullet" && a.style !== "numbered") {
          p.detachFromList();
          await ctx.sync();
        }
        return { ok: true, count: 1 };
      }
      case "insert": {
        const blocks = parseBlocks(a.text, a.style ?? "normal");
        if (!blocks.length) return { ok: false, error: "failed", detail: "Ingen tekst" };
        let anchor: Word.Paragraph | undefined;
        if (a.position === "before" || a.position === "after") {
          const p = this.paragraph(dc, session, a.paragraph);
          if (!("delete" in p)) return p;
          anchor = p;
        }
        let prev: Word.Paragraph | undefined;
        for (const b of blocks) {
          let p: Word.Paragraph;
          if (prev) p = prev.insertParagraph(b.inline.text, "After");
          else if (anchor) p = anchor.insertParagraph(b.inline.text, a.position === "before" ? "Before" : "After");
          else if (dc.isWholeDocument || !session.paras.length) p = ctx.document.body.insertParagraph(b.inline.text, a.position === "start" ? "Start" : "End");
          else
            p =
              a.position === "start"
                ? session.paras[0].insertParagraph(b.inline.text, "Before")
                : session.paras[session.paras.length - 1].insertParagraph(b.inline.text, "After");
          await this.formatNew(ctx, p, b);
          prev = p;
        }
        return { ok: true, count: blocks.length };
      }
      case "move": {
        const src = this.paragraph(dc, session, a.paragraph);
        if (!("delete" in src)) return src;
        const dst = this.paragraph(dc, session, a.to);
        if (!("delete" in dst)) return dst;
        // Hele afsnittet (inkl. afsnitsmærket) flyttes som OOXML, så typografi, lister og
        // inline-formatering følger med. Originalen slettes først, når indsættelsen er gået igennem.
        const xml = src.getRange("Whole").getOoxml();
        await ctx.sync();
        dst.getRange("Whole").insertOoxml(xml.value, a.position === "before" ? "Before" : "After");
        await ctx.sync();
        src.delete();
        await ctx.sync();
        session.gone.add(a.paragraph);
        return { ok: true, count: 1 };
      }
      default:
        return { ok: false, error: "failed", detail: `${a.action} findes ikke i Word` };
    }
  }
}
