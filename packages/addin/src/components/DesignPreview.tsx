import { useMemo, useState } from "react";
import { Body1, Button, Caption1, Divider, MessageBar, MessageBarBody, Switch, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowUndo16Regular, Play16Filled } from "@fluentui/react-icons";
import { DESIGN_LAYOUT_NAMES, type DesignPlan, type DesignSlide, type DocContext } from "@miai/shared";
import { kpiBoxes, resolvePalette, SLIDE_H, SLIDE_W, templateFor, type Palette, type Template } from "../deck/templates";
import type { DesignOutcome, HostAdapter } from "../hosts/types";
import { useI18n } from "../i18n";
import { ui } from "../theme";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  muted: { color: tokens.colorNeutralForeground3 },
  swatches: { display: "flex", flexWrap: "wrap", gap: tokens.spacingHorizontalS },
  swatch: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS },
  chip: {
    width: ui.iconTileSmall,
    height: ui.iconTileSmall,
    borderRadius: tokens.borderRadiusCircular,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
  },
  grid: { display: "grid", gridTemplateColumns: `repeat(auto-fill, minmax(${ui.thumbMin}, 1fr))`, gap: tokens.spacingHorizontalS },
  thumb: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXXS, minWidth: 0 },
  svg: { width: "100%", height: "auto", borderRadius: tokens.borderRadiusMedium, boxShadow: tokens.shadow4 },
  bar: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: tokens.spacingHorizontalS },
});

const S = 100; // tommer → viewBox-enheder
const pt = (size: number) => (size / 72) * S;
/** Skrifttypen med en reserve af samme slags, hvis den ikke findes på maskinen (fx Aptos uden Microsoft 365). */
const stack = (font: string) => `"${font}", ${/Georgia|Cambria|Garamond/.test(font) ? "Georgia, serif" : '"Segoe UI", Helvetica, Arial, sans-serif'}`;
const clip = (s: string, n: number) => (s.length > n ? `${s.slice(0, n - 1)}…` : s);

/** Tekstlinjer i en boks: groft ombrudt efter boksens bredde og skriftstørrelsen. */
function lines(text: string, w: number, size: number, max: number): string[] {
  const perLine = Math.max(8, Math.floor((w * 72) / (size * 0.52)));
  const out: string[] = [];
  for (const para of text.split("\n")) {
    let cur = "";
    for (const word of para.split(" ")) {
      if ((cur + " " + word).trim().length > perLine) {
        out.push(cur.trim());
        cur = word;
      } else cur = `${cur} ${word}`;
    }
    out.push(cur.trim());
  }
  return out.length > max ? [...out.slice(0, max - 1), clip(out[max - 1], perLine)] : out;
}

/** Miniature af ét slide, tegnet med samme skabelon som den .pptx, der indsættes. */
function SlideThumb({ plan, slide, pal, t, label }: { plan: DesignPlan; slide: DesignSlide; pal: Palette; t: Template; label: string }) {
  const fontOf = (f: "heading" | "body") => stack(f === "heading" ? plan.fonts.heading : plan.fonts.body);
  const textOf: Record<string, string | undefined> = {};
  switch (slide.layout) {
    case "title":
    case "section":
    case "closing":
      textOf.title = slide.title;
      textOf.subtitle = slide.subtitle;
      break;
    case "bullets":
      textOf.title = slide.title;
      textOf.body = slide.bullets.join("\n");
      break;
    case "twoColumn":
      Object.assign(textOf, {
        title: slide.title,
        leftHead: slide.left.heading,
        left: slide.left.bullets.join("\n"),
        rightHead: slide.right.heading,
        right: slide.right.bullets.join("\n"),
      });
      break;
    case "quote":
      textOf.quote = `“${slide.quote}”`;
      textOf.attribution = slide.attribution && `— ${slide.attribution}`;
      break;
    case "kpi":
      textOf.title = slide.title;
      break;
  }
  return (
    <svg className={useStyles().svg} viewBox={`0 0 ${SLIDE_W * S} ${SLIDE_H * S}`} role="img" aria-label={label}>
      <rect width={SLIDE_W * S} height={SLIDE_H * S} fill={pal[t.background]} />
      {t.shapes.map((r, i) => (
        <rect key={i} x={r.x * S} y={r.y * S} width={r.w * S} height={r.h * S} fill={pal[r.fill]} />
      ))}
      {t.placeholders.map((p) => {
        const value = textOf[p.name];
        if (!value) return null;
        const size = pt(p.size);
        const ls = lines(value, p.w, p.size, Math.max(1, Math.floor((p.h * 72) / (p.size * 1.25))));
        const height = ls.length * size * 1.2;
        const y0 = p.valign === "bottom" ? p.y * S + p.h * S - height : p.valign === "middle" ? p.y * S + (p.h * S - height) / 2 : p.y * S;
        const x = p.align === "center" ? (p.x + p.w / 2) * S : p.x * S + (p.bullets ? size * 0.9 : 0);
        return (
          <g key={p.name}>
            {ls.map((l, i) => (
              <g key={i}>
                {p.bullets && <circle cx={p.x * S + size * 0.3} cy={y0 + i * size * 1.2 + size * 0.65} r={size * 0.14} fill={pal.accent} />}
                <text
                  x={x}
                  y={y0 + (i + 1) * size * 1.2 - size * 0.25}
                  fontFamily={fontOf(p.font)}
                  fontSize={size}
                  fontWeight={p.bold ? 700 : 400}
                  fontStyle={p.italic ? "italic" : "normal"}
                  fill={pal[p.color]}
                  textAnchor={p.align === "center" ? "middle" : "start"}
                >
                  {l}
                </text>
              </g>
            ))}
          </g>
        );
      })}
      {slide.layout === "kpi" &&
        kpiBoxes(slide.kpis.length).map((b, i) => (
          <g key={i}>
            <text
              x={b.x * S}
              y={(b.valueY + 1.3) * S}
              fontFamily={stack(plan.fonts.heading)}
              fontSize={pt(slide.kpis.length > 3 ? 44 : 54)}
              fontWeight={700}
              fill={pal.accent}
            >
              {slide.kpis[i].value}
            </text>
            <text x={b.x * S} y={(b.labelY + 0.3) * S} fontFamily={stack(plan.fonts.body)} fontSize={pt(18)} fill={pal.muted}>
              {clip(slide.kpis[i].label, 28)}
            </text>
          </g>
        ))}
    </svg>
  );
}

type Phase = { kind: "review" } | { kind: "running" } | { kind: "done"; outcome: DesignOutcome } | { kind: "undone" };

/**
 * Designforslaget (ADR-0014): navn og begrundelse, palet, skrifttyper og en miniature af hvert
 * slide. "Opret" sætter masteren og slides ind; de nuværende slides kan erstattes og lægges
 * tilbage med Fortryd.
 */
export function DesignPreview({ plan, ctx, adapter }: { plan: DesignPlan; ctx: DocContext; adapter: HostAdapter }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const pal = useMemo(() => resolvePalette(plan), [plan]);
  const existing = ctx.slides?.length ?? 0;
  const [replace, setReplace] = useState(false);
  const [phase, setPhase] = useState<Phase>({ kind: "review" });
  const [error, setError] = useState<string>();

  const create = async () => {
    setError(undefined);
    setPhase({ kind: "running" });
    try {
      setPhase({ kind: "done", outcome: await adapter.applyDesign!(plan, { replace, lang }) });
    } catch (err) {
      setError((err as Error).message);
      setPhase({ kind: "review" });
    }
  };
  const undo = async () => {
    if (phase.kind !== "done") return;
    try {
      await phase.outcome.undo();
      setPhase({ kind: "undone" });
    } catch (err) {
      setError((err as Error).message);
    }
  };

  const roles: [keyof Palette, string][] = [
    ["primary", t("designPrimary")],
    ["accent", t("designAccent")],
    ["background", t("designBackground")],
    ["text", t("designText")],
    ["muted", t("designMuted")],
  ];

  return (
    <div className={s.root}>
      <div>
        <Text weight="semibold">{plan.name}</Text>
        {plan.rationale && <Body1 block>{plan.rationale}</Body1>}
      </div>
      <div className={s.swatches} aria-label={t("designPalette")}>
        {roles.map(([role, label]) => (
          <span key={role} className={s.swatch}>
            <span className={s.chip} style={{ backgroundColor: pal[role] }} aria-hidden />
            <Caption1>
              {label} {pal[role]}
            </Caption1>
          </span>
        ))}
      </div>
      <Caption1 className={s.muted}>
        {t("designFonts")}: <span style={{ fontFamily: stack(plan.fonts.heading) }}>{plan.fonts.heading}</span> /{" "}
        <span style={{ fontFamily: stack(plan.fonts.body) }}>{plan.fonts.body}</span>
      </Caption1>
      <div className={s.grid}>
        {plan.slides.map((sl, i) => {
          const label = `${i + 1}. ${DESIGN_LAYOUT_NAMES[sl.layout][lang]}`;
          return (
            <div key={i} className={s.thumb}>
              <SlideThumb plan={plan} slide={sl} pal={pal} t={templateFor(plan.style, sl.layout)} label={label} />
              <Caption1 className={s.muted}>{label}</Caption1>
            </div>
          );
        })}
      </div>
      <Divider />
      {phase.kind === "review" || phase.kind === "running" ? (
        <>
          {existing > 0 && (
            <Switch
              label={existing === 1 ? t("designReplaceOne") : t("designReplace", { n: existing })}
              checked={replace}
              disabled={phase.kind === "running"}
              onChange={(_, d) => setReplace(d.checked)}
            />
          )}
          <Caption1 className={s.muted}>{replace ? t("designReplaceHint") : t("designAddHint")}</Caption1>
          <div className={s.bar}>
            <Button appearance="primary" size="small" icon={<Play16Filled />} disabled={phase.kind === "running"} onClick={() => void create()}>
              {t("designCreate", { n: plan.slides.length })}
            </Button>
          </div>
        </>
      ) : phase.kind === "done" ? (
        <>
          <MessageBar intent={phase.outcome.keptOld ? "warning" : "success"}>
            <MessageBarBody>
              {t("designDone", { n: phase.outcome.inserted })}{" "}
              {phase.outcome.keptOld ? t("designKeptOld") : phase.outcome.removed ? t("designRemoved", { n: phase.outcome.removed }) : ""}
            </MessageBarBody>
          </MessageBar>
          <Caption1 className={s.muted}>{t("designAfter")}</Caption1>
          <div className={s.bar}>
            <Button size="small" icon={<ArrowUndo16Regular />} onClick={() => void undo()}>
              {t("undo")}
            </Button>
          </div>
        </>
      ) : (
        <MessageBar intent="success">
          <MessageBarBody>{t("designUndone")}</MessageBarBody>
        </MessageBar>
      )}
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}
