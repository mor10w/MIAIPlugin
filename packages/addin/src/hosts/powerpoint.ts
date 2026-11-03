import { applyEdits, diagramBounds, layoutDiagram, routeEdges, type ApplyMode, type DiagramResult, type DocContext, type GrammarResult, type SlidesResult } from "@miai/shared";
import { graphicToPng } from "../image";
import { mdToPlain } from "../markdown";
import { extractGraphic } from "../svg";
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
    if (mode === "newSlides" || mode === "diagram") return has("1.4");
    if (mode === "image") return has("1.8"); // billedfyldning kræver PowerPointApi 1.8
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

  /**
   * Indsætter diagrammet som native figurer på den valgte slide (ADR-0009), så brugeren kan
   * klikke på, flytte og redigere hver del.
   *
   * PowerPoint-API'et kan hverken vende linjer (negative mål afvises) eller sætte pilehoveder,
   * så kanterne tegnes som vinkelrette ruter fra `routeEdges` – kun vandrette og lodrette
   * segmenter – med en lille chevron som pilehoved. Med PowerPointApi 1.8 grupperes det hele.
   */
  async applyDiagram(result: DiagramResult): Promise<ApplyOutcome> {
    if (!has("1.4")) throw new Error("Din PowerPoint-version understøtter ikke figurer fra add-ins (kræver PowerPointApi 1.4).");
    const L = layoutDiagram(result);
    const routes = routeEdges(result, L);
    const B = diagramBounds(L, routes);
    // Skalér til slidens indholdsområde (højst 1,25x, så små diagrammer ikke bliver kæmpestore), og centrér.
    const s = Math.min(1.25, (W - 80) / B.width, (H - 80) / B.height);
    const ox = (W - B.width * s) / 2;
    const oy = (H - B.height * s) / 2;
    const X = (x: number) => ox + (x - B.x) * s;
    const Y = (y: number) => oy + (y - B.y) * s;
    const fontSize = Math.max(9, Math.min(20, Math.round(13 * s)));

    await PowerPoint.run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items");
      await ctx.sync();
      if (!sel.items.length) throw new Error("Vælg en slide, og prøv igen.");
      const shapes = sel.items[0].shapes;
      const added: PowerPoint.Shape[] = [];

      // Kanter først, så knuderne ligger ovenpå.
      for (const r of routes) {
        for (let i = 1; i < r.points.length; i++) {
          const p = r.points[i - 1];
          const q = r.points[i];
          if (p.x === q.x && p.y === q.y) continue;
          added.push(
            shapes.addLine(PowerPoint.ConnectorType.straight, {
              left: X(Math.min(p.x, q.x)),
              top: Y(Math.min(p.y, q.y)),
              width: Math.abs(q.x - p.x) * s,
              height: Math.abs(q.y - p.y) * s,
            }),
          );
        }
        // Pilehoved: alle kanter ender i målets venstre side, så en chevron (peger mod højre
        // uden rotation) passer altid.
        const end = r.points[r.points.length - 1];
        const a = Math.max(6, 8 * s);
        added.push(shapes.addGeometricShape(PowerPoint.GeometricShapeType.chevron, { left: X(end.x) - a, top: Y(end.y) - a / 2, width: a, height: a }));
      }

      for (const n of L.nodes) {
        const sh = shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, { left: X(n.x), top: Y(n.y), width: n.w * s, height: n.h * s });
        sh.textFrame.textRange.text = n.label;
        sh.textFrame.textRange.font.size = fontSize;
        sh.textFrame.wordWrap = true;
        added.push(sh);
      }
      await ctx.sync();

      if (has("1.8") && added.length > 1) {
        shapes.addGroup(added);
        await ctx.sync();
      }
    });
    const edges = routes.length;
    return { message: `${result.nodes.length} figurer og ${edges} forbindelser er placeret på den valgte slide${has("1.8") ? " som én gruppe" : ""} – klik på dem for at redigere.` };
  }

  /** Indsætter den første ```mermaid- eller ```svg-blok i teksten som billedfyldt rektangel på den valgte slide. */
  async applyImage(text: string): Promise<ApplyOutcome> {
    if (!has("1.8")) throw new Error("Din PowerPoint-version understøtter ikke billedfigurer fra add-ins (kræver PowerPointApi 1.8).");
    const block = extractGraphic(text);
    if (!block) throw new Error("Svaret indeholder ikke et diagram (mermaid eller svg).");
    const png = await graphicToPng(block);
    // Skalér til slide med margin, bevar forholdsforhold, og centrer.
    const maxW = W - 80;
    const maxH = H - 80;
    const s = Math.min(maxW / png.width, maxH / png.height);
    const w = Math.max(40, png.width * s);
    const h = Math.max(24, png.height * s);
    await PowerPoint.run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items");
      await ctx.sync();
      if (!sel.items.length) throw new Error("Vælg en slide, og prøv igen.");
      const sh = sel.items[0].shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: (W - w) / 2,
        top: (H - h) / 2,
        width: w,
        height: h,
      });
      sh.fill.setImage(png.base64);
      sh.lineFormat.visible = false; // ellers får billedet rektanglets kant
      sh.name = "MIAI-figur";
      await ctx.sync();
    });
    return { message: "Diagrammet er placeret på den valgte slide – klik på det for at flytte eller redigere." };
  }
}
