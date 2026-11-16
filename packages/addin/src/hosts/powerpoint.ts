import { runPowerPointScript } from "../script";
import { selectedText, wordCount } from "./selection";
import type { AgentOptions, SelectionInfo } from "./types";
import type { DocRef } from "../refs";
import {
  applyEdits,
  minimalEdit,
  parseBlocks,
  type AgentAction,
  type DesignPlan,
  diagramBounds,
  layoutDiagram,
  routeEdges,
  type ApplyMode,
  type DiagramResult,
  type DocContext,
  type GrammarResult,
  type SlidesResult,
} from "@miai/shared";
import { graphicToPng } from "../image";
import { mdToPlain } from "../markdown";
import { extractGraphic } from "../svg";
import {
  MAX_CONTEXT_CHARS,
  truncate,
  type AgentOutcome,
  type AgentStepResult,
  type ApplyOutcome,
  type ContextOptions,
  type DesignOutcome,
  type HostAdapter,
  type TextFormat,
} from "./types";
import { buildDeck } from "../deck/buildDeck";

const has = (v: string) => Office.context.requirements.isSetSupported("PowerPointApi", v);

function getSelectedText(): Promise<string> {
  return new Promise((resolve) =>
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? String(r.value ?? "") : ""),
    ),
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

/**
 * Figurtyper, der kan have tekst. For de andre (billede, tabel, diagram, gruppe …) afviser
 * PowerPoint `textFrame` med InvalidArgument – ved næste sync, så én sådan figur får hele den
 * samlede læsning til at fejle.
 */
const TEXT_SHAPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]);

/** Teksten i en figur; `range` kan skrives i. Pladsholdere får altid en tekst ("" når tom). */
interface ShapeText {
  text?: string;
  range?: PowerPoint.TextRange;
}

/**
 * Læser teksten i figurerne (`type` skal være indlæst). Kun figurer, der kan have tekst, spørges;
 * afviser PowerPoint alligevel én (fx en pladsholder, der rummer et billede), læses de én ad
 * gangen, og den, der fejler, springes over.
 */
async function readShapeTexts(ctx: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]): Promise<ShapeText[]> {
  const out: ShapeText[] = shapes.map(() => ({}));
  const candidates = shapes.flatMap((sh, i) => (TEXT_SHAPES.has(String(sh.type)) ? [i] : []));
  const read = async (idx: number[]) => {
    const frames = idx.map((i) => {
      const tf = shapes[i].textFrame;
      tf.load("hasText");
      return tf;
    });
    await ctx.sync();
    const ranges = frames.map((tf, k) => (tf.hasText || shapes[idx[k]].type === PowerPoint.ShapeType.placeholder ? tf.textRange : undefined));
    ranges.forEach((r) => r?.load("text"));
    await ctx.sync();
    idx.forEach((i, k) => (out[i] = ranges[k] ? { text: ranges[k]!.text, range: ranges[k] } : {}));
  };
  try {
    await read(candidates);
  } catch {
    for (const i of candidates) {
      try {
        await read([i]);
      } catch {
        out[i] = {};
      }
    }
  }
  return out;
}

export class PowerPointAdapter implements HostAdapter {
  readonly host = "powerpoint" as const;

  /** Antal markerede ord, eller null når intet er markeret (#26). */
  async selectionInfo(): Promise<SelectionInfo | null> {
    const words = wordCount(await selectedText());
    return words ? { kind: "text", words } : null;
  }

  /** Går til sliden (#22). Ændrer intet. */
  goTo(ref: DocRef): Promise<void> {
    return new Promise((resolve, reject) => {
      if (!ref.slide) return resolve();
      Office.context.document.goToByIdAsync(ref.slide, Office.GoToType.Index, (r) =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message ?? "Kunne ikke gå til sliden")),
      );
    });
  }

  supports(mode: ApplyMode): boolean {
    if (mode === "newSlides" || mode === "diagram") return has("1.4");
    if (mode === "image") return has("1.8"); // billedfyldning kræver PowerPointApi 1.8
    if (mode === "agent") return has("1.4");
    // Indsættelse af slides fra en fil (med master og layouts) kræver PowerPointApi 1.2.
    if (mode === "design") return has("1.2");
    return mode === "replace" || mode === "insertAfter";
  }

  async getContext(opts: ContextOptions): Promise<DocContext> {
    // Agenten ser hele præsentationen med layouts og figurer (PowerPointApi 1.4).
    if (opts.structure && has("1.4")) {
      const { slides, layouts } = await this.structure(opts.maxChars);
      const text = slides.map((sl, i) => `--- Slide ${i + 1} ---\n${sl.shapes.map((sh) => sh.text).join("\n")}`).join("\n\n");
      return { host: "powerpoint", ...truncate(text, opts.maxChars), isWholeDocument: true, slides, layouts };
    }
    const text = await getSelectedText();
    if (text.trim() || !opts.wholeDocumentFallback || !has("1.4")) return { host: "powerpoint", ...truncate(text, opts.maxChars), isWholeDocument: false };
    return { host: "powerpoint", ...truncate(await this.allText(), opts.maxChars), isWholeDocument: true };
  }

  /** Al tekst i præsentationen, slide for slide. */
  private allText(): Promise<string> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const shapeCols = slides.items.map((s) => {
        s.shapes.load("items/type");
        return s.shapes;
      });
      await ctx.sync();
      const texts: string[][] = [];
      for (const col of shapeCols) texts.push((await readShapeTexts(ctx, col.items)).flatMap((t) => (t.text ? [t.text] : [])));
      return texts.map((ts, i) => `--- Slide ${i + 1} ---\n${ts.join("\n")}`).join("\n\n");
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
    return {
      message: `${result.nodes.length} figurer og ${edges} forbindelser er placeret på den valgte slide${has("1.8") ? " som én gruppe" : ""} – klik på dem for at redigere.`,
    };
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

  // --- Agent-tilstand (#6) -----------------------------------------------------------------

  /** Pladsholderens rolle: titel, brødtekst/indhold, egen tekstboks eller figur uden tekst. */
  private kind(sh: PowerPoint.Shape, t: ShapeText, phType?: string): "title" | "body" | "text" | "other" {
    // En pladsholder uden tekstfelt (fx med et billede) er ikke noget, agenten kan skrive i.
    if (!t.range) return "other";
    if (sh.type === PowerPoint.ShapeType.placeholder) return this.role(sh, phType);
    return t.text ? "text" : "other";
  }

  /** En pladsholders rolle ud fra dens type (PowerPointApi 1.8) eller navn. */
  private role(sh: PowerPoint.Shape, phType?: string): "title" | "body" {
    const name = (phType ?? sh.name).toLowerCase();
    return /title|titel/.test(name) && !/subtitle|undertitel/.test(name) ? "title" : "body";
  }

  /** Slides med layout og figurer, plus præsentationens layouts (fra første slidemaster). */
  private async structure(maxChars = MAX_CONTEXT_CHARS): Promise<{ slides: NonNullable<DocContext["slides"]>; layouts: string[] }> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id,items/layout/name");
      const masters = ctx.presentation.slideMasters;
      masters.load("items/layouts/items/name");
      await ctx.sync();
      const shapeCols = slides.items.map((sl) => {
        sl.shapes.load("items/name,items/type");
        return sl.shapes;
      });
      await ctx.sync();
      const phs = shapeCols.map((col) =>
        col.items.map((sh) => {
          const ph = has("1.8") && sh.type === PowerPoint.ShapeType.placeholder ? sh.placeholderFormat : undefined;
          ph?.load("type");
          return ph;
        }),
      );
      await ctx.sync();
      const texts: ShapeText[][] = [];
      for (const col of shapeCols) texts.push(await readShapeTexts(ctx, col.items));
      let chars = 0;
      const out = slides.items.map((sl, i) => ({
        layout: sl.layout.name,
        shapes: shapeCols[i].items.map((sh, j) => {
          const text = texts[i][j].text ?? "";
          chars += text.length;
          return { name: sh.name, kind: this.kind(sh, texts[i][j], phs[i][j]?.type), text: chars > maxChars ? "" : text };
        }),
      }));
      const layouts = [...new Set(masters.items.flatMap((m) => m.layouts.items.map((l) => l.name)))];
      return { slides: out, layouts };
    });
  }

  /**
   * Agent-tilstand (#6): ændrer præsentationen og bevarer skabelonens design:
   *  - erstat ændrer kun de ord, der er forskellige (resten af tekstens formatering bevares);
   *  - ny tekst i en figur får figurens egen skrifttype, størrelse, farve og punkttegn;
   *  - nye slides oprettes fra præsentationens layouts og udfyldes i pladsholderne – ingen egne
   *    tekstbokse, farver eller skrifttyper.
   * Fortryd (PowerPointApi 1.8): de berørte slides gemmes, før de ændres, og lægges tilbage
   * med deres formatering; tilføjede slides slettes, og flyttede flyttes tilbage. Ældre
   * PowerPoint kan kun gendanne teksten, og sletning af slides kræver 1.8.
   */
  async applyAgent(actions: AgentAction[], dc: DocContext, opts: AgentOptions = {}): Promise<AgentOutcome> {
    // Kodetilstand (#23): et script står alene i planen og kører i sandkassen med sit eget øjebliksbillede.
    const script = actions.find((a) => a.action === "runScript");
    if (script?.action === "runScript") {
      if (!opts.scripts) throw new Error("Kodetilstanden er ikke slået til (Indstillinger → Avanceret).");
      if (actions.length > 1) throw new Error("Et script skal stå alene i planen.");
      return runPowerPointScript(script);
    }
    const full = has("1.8");
    return PowerPoint.run(async (ctx) => {
      const slideCol = ctx.presentation.slides;
      slideCol.load("items/id,items/layout/id");
      const masters = ctx.presentation.slideMasters;
      masters.load("items/id,items/layouts/items/name,items/layouts/items/id");
      await ctx.sync();
      const slides = [...slideCol.items];
      const ids = slides.map((s) => s.id);
      const shapeCols = slides.map((sl) => {
        sl.shapes.load("items/name,items/type");
        return sl.shapes;
      });
      await ctx.sync();
      const textOf: (PowerPoint.TextRange | undefined)[][] = [];
      for (const col of shapeCols) textOf.push((await readShapeTexts(ctx, col.items)).map((t) => t.range));
      /** Figurernes tekst, som den er nu – opdateres lokalt, når agenten ændrer den. */
      const current = textOf.map((trs) => trs.map((tr) => tr?.text ?? ""));
      // Er præsentationen ændret siden planen? (tjekkes én gang, før første handling)
      const stale = new Set<number>();
      dc.slides?.forEach((sl, i) => {
        const now = textOf[i]?.map((t) => t?.text ?? "") ?? [];
        if (sl.shapes.length !== now.length || sl.shapes.some((sh, j) => sh.text !== "" && sh.text !== now[j])) stale.add(i + 1);
      });

      // Slides, handlingerne ændrer, gemmes til fortryd før første ændring.
      const touched = new Set<number>();
      for (const a of actions) {
        if (a.action === "setText") touched.add(Number(a.shape.split(".")[0]));
        if (a.action === "deleteSlide") touched.add(a.slide);
        if (a.action === "replace")
          dc.slides?.forEach((sl, i) => {
            if (sl.shapes.some((sh) => sh.text.includes(a.find))) touched.add(i + 1);
          });
      }
      const exported = full ? [...touched].map((n) => ({ n, id: ids[n - 1], data: slides[n - 1]?.exportAsBase64() })) : [];
      const textBackup = !full ? [...touched].flatMap((n) => (textOf[n - 1] ?? []).map((tr, j) => ({ n, j, text: tr?.text }))) : [];
      await ctx.sync();

      const results: AgentStepResult[] = [];
      const added: string[] = [];
      const moved: { id: string; from: number }[] = [];
      const gone = new Set<number>();
      const slideRef = (n: number): PowerPoint.Slide | AgentStepResult => {
        if (stale.has(n)) return { ok: false, error: "changed" };
        if (gone.has(n)) return { ok: false, error: "failed", detail: `Slide ${n} er slettet af en tidligere handling` };
        return slides[n - 1] ?? { ok: false, error: "failed", detail: `Slide ${n} findes ikke` };
      };

      for (const a of actions) {
        try {
          switch (a.action) {
            case "replace": {
              const { before, oldCore, newCore } = minimalEdit(a.find, a.replace);
              let count = 0;
              for (const [i, trs] of textOf.entries()) {
                if (stale.has(i + 1) || gone.has(i + 1)) continue;
                for (const [j, tr] of trs.entries()) {
                  if (!tr || (count && !a.all)) continue;
                  const txt = current[i][j];
                  const at = txt.indexOf(a.find);
                  if (at < 0 || !oldCore) continue;
                  // Kun den ændrede del skiftes ud – formateringen omkring den bevares.
                  tr.getSubstring(at + before.length, oldCore.length).text = newCore;
                  current[i][j] = txt.slice(0, at + before.length) + newCore + txt.slice(at + before.length + oldCore.length);
                  count++;
                }
              }
              await ctx.sync();
              results.push(count ? { ok: true, count } : { ok: false, error: "not-found" });
              break;
            }
            case "setText": {
              const [sn, shn] = a.shape.split(".").map(Number);
              const sl = slideRef(sn);
              if (!("shapes" in sl)) {
                results.push(sl);
                break;
              }
              const tr = textOf[sn - 1]?.[shn - 1];
              if (!tr) {
                results.push({ ok: false, error: "failed", detail: `Figur ${a.shape} har ingen tekst` });
                break;
              }
              await this.writeText(ctx, tr, a.text);
              current[sn - 1][shn - 1] = parseBlocks(a.text)
                .map((b) => b.inline.text)
                .join("\n");
              results.push({ ok: true, count: 1 });
              break;
            }
            case "addSlide": {
              // Layout: det navngivne, ellers slidet før (så det nye følger samme skabelon).
              const all = masters.items.flatMap((m) => m.layouts.items.map((l) => ({ master: m.id, id: l.id, name: l.name })));
              const prev = slides[a.after - 1];
              const layout = (a.layout && all.find((l) => l.name === a.layout)) || all.find((l) => l.id === prev?.layout.id) || all[0];
              ctx.presentation.slides.add(layout ? { layoutId: layout.id, slideMasterId: layout.master } : undefined);
              await ctx.sync();
              const now = ctx.presentation.slides;
              now.load("items/id");
              await ctx.sync();
              const slide = now.items[now.items.length - 1];
              added.push(slide.id);
              if (full) slide.moveTo(a.after);
              slide.shapes.load("items/name,items/type");
              await ctx.sync();
              const phs = slide.shapes.items.filter((sh) => sh.type === PowerPoint.ShapeType.placeholder);
              const title = phs.find((sh) => this.role(sh) === "title") ?? phs[0];
              const body = phs.find((sh) => sh !== title);
              if (title) title.textFrame.textRange.text = a.title;
              if (body && a.body) await this.writeText(ctx, body.textFrame.textRange, a.body);
              await ctx.sync();
              const missing = !title || (a.body && !body);
              results.push(missing ? { ok: true, count: 1, detail: "Layoutet har ikke pladsholdere til al teksten" } : { ok: true, count: 1 });
              break;
            }
            case "deleteSlide": {
              const sl = slideRef(a.slide);
              if (!("shapes" in sl)) {
                results.push(sl);
                break;
              }
              // Uden PowerPointApi 1.8 kan et slettet slide ikke lægges tilbage – så slettes det ikke.
              if (!full) {
                results.push({ ok: false, error: "failed", detail: "Sletning af slides kræver en nyere PowerPoint (så det kan fortrydes)" });
                break;
              }
              sl.delete();
              await ctx.sync();
              gone.add(a.slide);
              results.push({ ok: true, count: 1 });
              break;
            }
            case "moveSlide": {
              const sl = slideRef(a.slide);
              if (!("shapes" in sl)) {
                results.push(sl);
                break;
              }
              if (!full) {
                results.push({ ok: false, error: "failed", detail: "Flytning af slides kræver PowerPointApi 1.8" });
                break;
              }
              sl.load("index");
              await ctx.sync();
              moved.push({ id: sl.id, from: sl.index });
              sl.moveTo(a.to - 1);
              await ctx.sync();
              results.push({ ok: true, count: 1 });
              break;
            }
            default:
              results.push({ ok: false, error: "failed", detail: `${a.action} findes ikke i PowerPoint` });
          }
        } catch (err) {
          results.push({ ok: false, error: "failed", detail: (err as Error).message });
        }
      }

      const exportedData = exported.map((e) => ({ n: e.n, id: e.id, base64: e.data?.value }));
      const undo = () =>
        PowerPoint.run(async (c) => {
          const col = c.presentation.slides;
          col.load("items/id");
          await c.sync();
          // Tilføjede slides slettes, flyttede flyttes tilbage (omvendt rækkefølge).
          for (const id of added) col.items.find((s) => s.id === id)?.delete();
          for (const m of [...moved].reverse()) col.items.find((s) => s.id === m.id)?.moveTo(m.from);
          await c.sync();
          if (full) {
            // Ændrede og slettede slides lægges tilbage fra kopien – med al formatering – på deres plads.
            for (const e of [...exportedData].sort((x, y) => x.n - y.n)) {
              if (!e.base64) continue;
              const cur = c.presentation.slides;
              cur.load("items/id");
              await c.sync();
              const before = e.n > 1 ? cur.items[e.n - 2]?.id : undefined;
              c.presentation.insertSlidesFromBase64(e.base64, {
                formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
                ...(before ? { targetSlideId: before } : {}),
              });
              cur.items.find((s) => s.id === e.id)?.delete();
              await c.sync();
            }
          } else {
            const cur = c.presentation.slides;
            cur.load("items/id");
            await c.sync();
            for (const b of textBackup) {
              const sl = cur.items.find((s) => s.id === ids[b.n - 1]);
              if (!sl || b.text === undefined) continue;
              sl.shapes.load("items");
              await c.sync();
              const sh = sl.shapes.items[b.j];
              if (sh) sh.textFrame.textRange.text = b.text;
            }
            await c.sync();
          }
        });
      return { results, undo, undoKind: full ? "slides" : "text" };
    });
  }

  /**
   * Skriver tekst i en figur: én linje pr. afsnit (punkttegn kommer fra figurens egen typografi,
   * så "- " fjernes), og fed/kursiv sættes på de markerede ord. Resten af formateringen er figurens.
   */
  private async writeText(ctx: PowerPoint.RequestContext, tr: PowerPoint.TextRange, text: string): Promise<void> {
    const blocks = parseBlocks(text, "normal");
    const plain = blocks.map((b) => b.inline.text).join("\n");
    tr.text = plain;
    await ctx.sync();
    let offset = 0;
    for (const b of blocks) {
      for (const sp of b.inline.spans) {
        const sub = tr.getSubstring(offset + sp.start, sp.end - sp.start);
        if (sp.bold) sub.font.bold = true;
        if (sp.italic) sub.font.italic = true;
      }
      offset += b.inline.text.length + 1;
    }
    await ctx.sync();
  }

  // --- Design: master og layout (ADR-0014) ---------------------------------------------------

  /**
   * Sætter designets slides – og dermed dets master og layouts – ind sidst i præsentationen
   * ("behold kildeformatering"). Med replace fjernes de nuværende slides bagefter; de gemmes
   * først (PowerPointApi 1.8), så fortryd kan lægge dem tilbage med al formatering.
   */
  async applyDesign(plan: DesignPlan, opts: { replace: boolean; lang: "da" | "en" }): Promise<DesignOutcome> {
    const base64 = await buildDeck(plan, opts.lang);
    const full = has("1.8");
    return PowerPoint.run(async (ctx) => {
      const col = ctx.presentation.slides;
      col.load("items/id");
      await ctx.sync();
      const before = col.items.map((s) => s.id);
      const removeOld = opts.replace && full && before.length > 0;
      const saved = removeOld ? col.items.map((s) => s.exportAsBase64()) : [];
      await ctx.sync();
      ctx.presentation.insertSlidesFromBase64(base64, {
        formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
        ...(before.length ? { targetSlideId: before[before.length - 1] } : {}),
      });
      await ctx.sync();
      const after = ctx.presentation.slides;
      after.load("items/id");
      await ctx.sync();
      const inserted = after.items.map((s) => s.id).filter((id) => !before.includes(id));
      if (removeOld) {
        for (const s of after.items) if (before.includes(s.id)) s.delete();
        await ctx.sync();
      }
      const oldSlides = saved.map((r) => r.value);
      const undo = () =>
        PowerPoint.run(async (c) => {
          const cur = c.presentation.slides;
          cur.load("items/id");
          await c.sync();
          // De gamle slides lægges tilbage forrest i den oprindelige rækkefølge …
          const known = new Set(cur.items.map((x) => x.id));
          let target: string | undefined;
          for (const b64 of oldSlides) {
            c.presentation.insertSlidesFromBase64(b64, {
              formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
              ...(target ? { targetSlideId: target } : {}),
            });
            await c.sync();
            const now = c.presentation.slides;
            now.load("items/id");
            await c.sync();
            target = now.items.find((x) => !known.has(x.id))?.id ?? target;
            if (target) known.add(target);
          }
          // … og designets slides fjernes.
          for (const s of cur.items) if (inserted.includes(s.id)) s.delete();
          await c.sync();
        });
      return { inserted: inserted.length, removed: removeOld ? before.length : 0, keptOld: opts.replace && !removeOld && before.length > 0, undo };
    });
  }
}
