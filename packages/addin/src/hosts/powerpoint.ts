import { applyEdits, type ApplyMode, type DocContext, type GrammarResult, type SlidesResult } from "@miai/shared";
import { mdToPlain } from "../markdown";
import { truncate, type ApplyOutcome, type ContextOptions, type HostAdapter, type TextFormat } from "./types";

const has = (v: string) => Office.context.requirements.isSetSupported("PowerPointApi", v);

function getSelectedText(): Promise<string> {
  return new Promise((resolve) =>
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? String(r.value ?? "") : "")),
  );
}

function setSelectedText(text: string): Promise<void> {
  return new Promise((resolve, reject) =>
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message ?? "Kunne ikke indsætte tekst. Markér en tekstboks.")),
    ),
  );
}

// Standard 16:9-slide i punkter.
const W = 960;
const H = 540;

export class PowerPointAdapter implements HostAdapter {
  readonly host = "powerpoint" as const;

  supports(mode: ApplyMode): boolean {
    if (mode === "newSlides") return has("1.4");
    return mode === "replace" || mode === "insertAfter";
  }

  async getContext(opts: ContextOptions): Promise<DocContext> {
    const text = await getSelectedText();
    if (text.trim() || !opts.wholeDocumentFallback || !has("1.4")) return { host: "powerpoint", text: truncate(text), isWholeDocument: false };
    return { host: "powerpoint", text: truncate(await this.allText()), isWholeDocument: true };
  }

  /** Al tekst i præsentationen, slide for slide. */
  private allText(): Promise<string> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const shapeCols = slides.items.map((s) => {
        s.shapes.load("items/textFrame/hasText");
        return s.shapes;
      });
      await ctx.sync();
      const ranges = shapeCols.map((col) =>
        col.items
          .filter((sh) => sh.textFrame.hasText)
          .map((sh) => {
            sh.textFrame.textRange.load("text");
            return sh.textFrame.textRange;
          }),
      );
      await ctx.sync();
      return ranges.map((rs, i) => `--- Slide ${i + 1} ---\n${rs.map((r) => r.text).join("\n")}`).join("\n\n");
    }).catch(() => "");
  }

  async applyText(text: string, format: TextFormat, mode: ApplyMode, dc: DocContext): Promise<ApplyOutcome> {
    const plain = format === "markdown" ? mdToPlain(text) : text;
    await setSelectedText(mode === "insertAfter" && dc.text && !dc.isWholeDocument ? `${dc.text}\n${plain}` : plain);
    return {};
  }

  async applyGrammar(result: GrammarResult, _mode: ApplyMode, dc: DocContext): Promise<ApplyOutcome> {
    const { text, skipped } = applyEdits(dc.text, result.edits);
    await setSelectedText(text);
    const n = result.edits.length - skipped;
    return { message: `${n} af ${result.edits.length} rettelser anvendt (PowerPoint har ikke sporede ændringer).`, warning: skipped > 0 };
  }

  /**
   * Opretter slides sidst i præsentationen. Udfylder layoutets pladsholdere (titel + indhold), hvis de findes,
   * ellers tilføjes tekstbokse. Skal verificeres mod forskellige skabeloner i fase 0.
   */
  async applySlides(result: SlidesResult): Promise<ApplyOutcome> {
    if (!has("1.4")) throw new Error("Din PowerPoint-version understøtter ikke oprettelse af slides fra add-ins (kræver PowerPointApi 1.4).");
    await PowerPoint.run(async (ctx) => {
      for (const s of result.slides) {
        ctx.presentation.slides.add();
        await ctx.sync();
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const slide = slides.items[slides.items.length - 1];
        slide.shapes.load("items/type");
        await ctx.sync();
        const textShapes = slide.shapes.items.filter((sh) => sh.type === PowerPoint.ShapeType.placeholder);
        const bullets = s.bullets.join("\n");
        if (textShapes.length >= 2) {
          textShapes[0].textFrame.textRange.text = s.title;
          textShapes[1].textFrame.textRange.text = bullets;
        } else {
          textShapes.forEach((sh) => sh.delete());
          const title = slide.shapes.addTextBox(s.title, { left: 48, top: 32, width: W - 96, height: 80 });
          title.textFrame.textRange.font.size = 32;
          title.textFrame.textRange.font.bold = true;
          const body = slide.shapes.addTextBox(s.bullets.map((b) => `• ${b}`).join("\n"), { left: 48, top: 130, width: W - 96, height: H - 170 });
          body.textFrame.textRange.font.size = 20;
        }
        await ctx.sync();
      }
    });
    return { message: `${result.slides.length} slides oprettet sidst i præsentationen.` };
  }
}
