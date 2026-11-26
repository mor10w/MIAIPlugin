import { useEffect, useMemo, useRef, useState } from "react";
import {
  Badge,
  Body1,
  Button,
  Caption1,
  Checkbox,
  Divider,
  Link,
  MessageBar,
  MessageBarBody,
  Switch,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from "@fluentui/react-components";
import { ArrowUndo16Regular, CheckmarkCircle16Regular, ErrorCircle16Regular, Play16Filled } from "@fluentui/react-icons";
import { ui } from "../theme";
import { targetOf } from "../refs";
import { checkPlan, isDestructive, parseBlocks, writesCells, type AgentAction, type AgentPlan, type DocContext, type ExportSheetRequest } from "@miai/shared";
import type { AgentOutcome, HostAdapter, Overwrite } from "../hosts/types";
import { useI18n, type T } from "../i18n";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS, margin: 0, padding: 0, listStyleType: "none" },
  item: {
    display: "flex",
    alignItems: "flex-start",
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalXS}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
  },
  itemText: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXXS, minWidth: 0, flex: 1, paddingTop: tokens.spacingVerticalXS },
  what: { overflowWrap: "anywhere" },
  muted: { color: tokens.colorNeutralForeground3 },
  show: { fontSize: tokens.fontSizeBase200, alignSelf: "flex-start" },
  code: {
    marginTop: tokens.spacingVerticalXXS,
    "& summary": { cursor: "pointer", color: tokens.colorBrandForeground1, fontSize: tokens.fontSizeBase200 },
    "& pre": {
      margin: `${tokens.spacingVerticalXS} 0 0`,
      padding: tokens.spacingHorizontalS,
      maxHeight: ui.codeMaxHeight,
      overflow: "auto",
      fontSize: tokens.fontSizeBase100,
      backgroundColor: tokens.colorNeutralBackground3,
      borderRadius: tokens.borderRadiusMedium,
      whiteSpace: "pre-wrap",
    },
  },
  badges: { display: "flex", flexWrap: "wrap", gap: tokens.spacingHorizontalXS },
  status: { display: "inline-flex", alignItems: "center", gap: tokens.spacingHorizontalXXS },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  fail: { color: tokens.colorPaletteRedForeground1 },
  bar: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: tokens.spacingHorizontalS },
});

const clip = (s: string, n = 80) => (s.length > n ? `${s.slice(0, n - 1)}…` : s);

/** Afsnit ¶n som "¶3 “Priser…”", så brugeren kan se, hvilket afsnit der menes. */
const paraLabel = (n: number | undefined, ctx?: DocContext) => {
  const p = n !== undefined ? ctx?.paragraphs?.[n - 1] : undefined;
  return p ? `¶${n} “${clip(p.text, 40)}”` : `¶${n ?? "?"}`;
};

/** Menneskelig beskrivelse af én handling (da/en). */
export function describeAction(a: AgentAction, t: T, ctx?: DocContext): string {
  switch (a.action) {
    case "replace":
      return t(a.all ? "agentAct_replaceAll" : "agentAct_replace", { find: clip(a.find, 50), replace: clip(a.replace, 50) });
    case "insert": {
      const where =
        a.position === "start"
          ? t("agentPosStart")
          : a.position === "end"
            ? t("agentPosEnd")
            : t(a.position === "before" ? "agentPosBefore" : "agentPosAfter", { anchor: paraLabel(a.paragraph, ctx) });
      // Vis teksten uden Markdown-tegn (som den vil se ud i dokumentet).
      const plain = parseBlocks(a.text)
        .map((b) => b.inline.text)
        .join(" · ");
      return `${t("agentAct_insert", { text: clip(plain) })} – ${where}`;
    }
    case "delete":
      return a.paragraph !== undefined
        ? t("agentAct_deleteParagraph", { paragraph: paraLabel(a.paragraph, ctx) })
        : t("agentAct_delete", { find: clip(a.find ?? "") });
    case "move":
      return t("agentAct_move", {
        find: paraLabel(a.paragraph, ctx),
        where: t(a.position === "before" ? "agentBefore" : "agentAfter"),
        anchor: paraLabel(a.to, ctx),
      });
    case "setStyle":
      return t("agentAct_setStyle", { paragraph: paraLabel(a.paragraph, ctx), style: t(`agentStyle_${a.style}` as const) });
    case "comment":
      return t("agentAct_comment", { find: clip(a.find, 40), text: clip(a.text) });
    case "insertTable":
      return t("agentAct_insertTable", { rows: a.rows.length, cols: a.rows[0].length, header: clip(a.rows[0].join(" | "), 60) });
    case "addTableRow":
      return t("agentAct_addTableRow", { values: clip(a.values.join(" | "), 60), paragraph: paraLabel(a.paragraph, ctx) });
    case "deleteTableRow":
      return t("agentAct_deleteTableRow", { paragraph: paraLabel(a.paragraph, ctx) });
    case "setText": {
      const [sn, shn] = a.shape.split(".").map(Number);
      const sh = ctx?.slides?.[sn - 1]?.shapes[shn - 1];
      const plain = parseBlocks(a.text)
        .map((b) => b.inline.text)
        .join(" · ");
      return t("agentAct_setText", { shape: sh ? `${sh.name} (slide ${sn})` : a.shape, text: clip(plain) });
    }
    case "addSlide":
      return t("agentAct_addSlide", { title: clip(a.title, 50), after: a.after, layout: a.layout ?? "" }).replace(/ \(\)$/, "");
    case "deleteSlide":
      return t("agentAct_deleteSlide", { slide: a.slide, title: clip(ctx?.slides?.[a.slide - 1]?.shapes.find((s) => s.kind === "title")?.text ?? "", 40) });
    case "moveSlide":
      return t("agentAct_moveSlide", { slide: a.slide, to: a.to });
    case "appendRows":
      return t("agentAct_appendRows", { n: a.values.length, table: a.table });
    case "setValues":
      return t("agentAct_setValues", { n: a.values.reduce((n, r) => n + r.length, 0), address: a.address });
    case "setFormula":
      return t("agentAct_setFormula", { address: a.address, formula: clip(a.formula) });
    case "clear":
      return t("agentAct_clear", { address: a.address });
    case "deleteRows":
      return t("agentAct_deleteRows", { address: a.address });
    case "deleteEmptyRows":
      return t("agentAct_deleteEmptyRows", { target: a.target });
    case "formatTable":
      return t("agentAct_formatTable", { target: a.target, style: a.style.replace(/^TableStyle/, "") });
    case "setNumberFormat":
      return t("agentAct_setNumberFormat", { address: a.address, format: a.format });
    case "autofitColumns":
      return t("agentAct_autofitColumns", { address: a.address });
    case "sortRange":
      return t(a.order === "desc" ? "agentAct_sortDesc" : "agentAct_sortAsc", { target: a.target, by: a.by });
    case "conditionalFormat":
      return t("agentAct_conditionalFormat", { address: a.address, rule: t(`agentCf_${a.rule}` as const, { value: a.value ?? "" }) });
    case "addChart":
      return t("agentAct_addChart", { type: t(`agentChart_${a.type}` as const), source: a.source }) + (a.title ? ` “${clip(a.title, 40)}”` : "");
    case "readRange":
      return t("agentAct_readRange", { address: a.address });
    case "addPivotTable":
      return t("agentAct_addPivotTable", {
        source: a.source,
        rows: [...a.rows, ...a.columns].join(", "),
        values: a.values.map((v) => `${t(`agentPivot_${v.summarize}` as const)} ${v.field}`).join(", "),
        destination: a.destination,
      });
    case "setValidation":
      return a.list
        ? t("agentAct_validationList", { address: a.address, values: clip(a.list.join(", "), 80) })
        : t("agentAct_validationNumber", {
            address: a.address,
            kind: t(a.whole ? "agentValidationWhole" : "agentValidationDecimal"),
            range: a.min !== undefined && a.max !== undefined ? `${a.min}–${a.max}` : a.min !== undefined ? `≥ ${a.min}` : `≤ ${a.max}`,
          });
    case "addSheet":
      return t("agentAct_addSheet", { name: a.name });
    case "renameSheet":
      return t("agentAct_renameSheet", { sheet: a.sheet, name: a.name });
    case "freezeRows":
      return a.rows === 0 ? t("agentAct_unfreeze") : t("agentAct_freezeRows", { n: a.rows });
    case "runScript":
      return t("agentAct_runScript", { description: clip(a.description, 200) });
    case "saveSheet":
      return a.sheet ? t("agentAct_saveSheetNamed", { sheet: a.sheet, fileName: a.fileName }) : t("agentAct_saveSheet", { fileName: a.fileName });
  }
}

/** Advarsel før overskrivning (#75): "Overskriver 48 celler, heraf 12 formler"; ryd siger "Rydder …". */
export function overwriteLabel(a: AgentAction, n: Overwrite, t: T): string {
  const text =
    a.action === "clear"
      ? t(n.cells === 1 ? "agentClearFilledOne" : "agentClearFilled", { n: n.cells })
      : t(n.cells === 1 ? "agentOverwriteOne" : "agentOverwrite", { n: n.cells });
  return n.formulas > 0 ? t(n.formulas === 1 ? "agentWithFormulasOne" : "agentWithFormulas", { text, n: n.formulas }) : text;
}

/** Årsagen til, at en handling ikke kan udføres, som modellen skal have den (engelsk, som resten af prompten). */
const PROBLEM_HELP: Partial<Record<string, string>> = {
  "unsafe-script": "the script uses something the sandbox does not allow (e.g. fetch or another global that does not exist)",
  "with-script": "runScript must be the ONLY action in a plan; use the fixed actions instead of a script when they can do the job",
  "scripts-off": "runScript is not available; use only the listed actions",
  "wrong-host": "the action does not exist in this app",
  "bad-sheet": "the sheet does not exist (add it with addSheet first) or the name is taken",
  "bad-address": "the cell address is not valid",
  "bad-table": "there is no table with that name",
  "too-large": "too many cells in one action; split it",
  invalid: "the action is incomplete or contradicts itself",
};

type Phase = { kind: "review" } | { kind: "running" } | { kind: "done"; outcome: AgentOutcome; ran: number[] } | { kind: "undone" };

/**
 * Agentens plan (#6): plan → bekræft → udfør. Intet ændres, før brugeren har valgt handlinger
 * og trykket Udfør. Handlinger, der ikke kan udføres (tekst findes ikke osv.), vises fravalgt
 * med årsagen. Efter udførelsen kan hele sessionen fortrydes.
 */
export function AgentPlanView({
  plan,
  ctx,
  adapter,
  scripts = false,
  exportSheet,
  onRetry,
  onOutcome,
}: {
  plan: AgentPlan;
  ctx: DocContext;
  adapter: HostAdapter;
  /** Kodetilstanden er slået til og tilladt (ADR-0017). */
  scripts?: boolean;
  /** Eksport af et ark som .xlsx (tillæg ADR-0012); leveres kun med bridge. */
  exportSheet?: (data: ExportSheetRequest) => Promise<string>;
  /** Et script fejlede: lad modellen rette det med fejlbeskeden. */
  onRetry?: (retry: { error: string; code: string; partial?: boolean; blocked?: boolean }) => void;
  /** Planen er udført (eller fortrudt): en kort tekst til samtalens historik, så modellen kender udfaldet. */
  onOutcome?: (log: string) => void;
}) {
  const { t } = useI18n();
  const s = useStyles();
  const checked = useMemo(() => checkPlan(plan, ctx, { scripts, export: !!exportSheet }), [plan, ctx, scripts, exportSheet]);
  const [selected, setSelected] = useState<boolean[]>(() => checked.map((c) => !c.problem));
  const canTrack = adapter.agentTracking?.() ?? false;
  const [track, setTrack] = useState(canTrack);
  const [phase, setPhase] = useState<Phase>({ kind: "review" });
  const [error, setError] = useState<string>();
  const count = selected.filter(Boolean).length;

  // Advarsel før overskrivning (#75): adapteren læser arket (ændrer intet) og tæller, hvor mange
  // celler med indhold og formler hver skrivende handling rammer. Udfør venter, til tallene er der,
  // så ingen overskriver noget uden at have set advarslen. Kan arket ikke læses, vises ingen advarsel.
  const writers = useMemo(() => checked.flatMap((c, i) => (!c.problem && writesCells(c.action) ? [i] : [])), [checked]);
  const needsCount = !!adapter.overwrites && writers.length > 0;
  const [overwrites, setOverwrites] = useState<Record<number, Overwrite>>();
  const counted = useRef(false);
  useEffect(() => {
    if (!needsCount || counted.current) return;
    counted.current = true;
    adapter.overwrites!(
      writers.map((i) => checked[i].action),
      ctx,
    ).then(
      (r) => setOverwrites(Object.fromEntries(writers.flatMap((i, k) => (r[k] ? [[i, r[k]]] : [])))),
      () => setOverwrites({}),
    );
  }, [needsCount, adapter, writers, checked, ctx]);
  const counting = needsCount && !overwrites;

  // Markering af ændrede celler (#75): fjernes med knappen, ved fortryd, og når planen forsvinder
  // fra panelet (fx en ny plan), så farven ikke bliver stående uden en knap til at fjerne den.
  const [marked, setMarked] = useState(false);
  const [marks] = useState<{ unmark?: () => Promise<void> }>({});
  useEffect(() => () => void marks.unmark?.().catch(() => {}), [marks]);
  const unmarked = () => {
    marks.unmark = undefined;
    setMarked(false);
  };

  const execute = async () => {
    const ran = checked.flatMap((c, i) => (selected[i] && !c.problem ? [i] : []));
    setError(undefined);
    setPhase({ kind: "running" });
    try {
      const outcome = await adapter.applyAgent!(
        ran.map((i) => checked[i].action),
        ctx,
        // Markeringen kan også blive fjernet af en senere plan; så skal knappen her væk.
        { trackChanges: track, scripts, exportSheet, onUnmarked: unmarked },
      );
      setPhase({ kind: "done", outcome, ran });
      marks.unmark = outcome.unmark;
      setMarked(!!outcome.unmark);
      onOutcome?.(
        ran
          .map((i, k) => {
            const r = outcome.results[k];
            return `- ${describeAction(checked[i].action, t, ctx)}: ${r.ok ? "done" : `FAILED (${r.error ?? "failed"}${r.detail ? `: ${r.detail}` : ""})`}`;
          })
          .join("\n"),
      );
    } catch (err) {
      setError((err as Error).message);
      setPhase({ kind: "review" });
    }
  };

  const undo = async () => {
    if (phase.kind !== "done") return;
    setError(undefined);
    try {
      await phase.outcome.undo();
      unmarked();
      setPhase({ kind: "undone" });
      onOutcome?.("The user undid all of these changes; the document is as it was before the plan.");
    } catch (err) {
      setError((err as Error).message);
    }
  };

  const unmark = async () => {
    setError(undefined);
    try {
      await marks.unmark?.();
      unmarked();
    } catch (err) {
      setError((err as Error).message);
    }
  };

  /**
   * Faste handlinger, der fejlede (#20): fejlene og handlingerne kan sendes til modellen, som
   * foreslår en rettelse. Det, der lykkedes, er udført og skal ikke gentages.
   */
  const failedRetry = (() => {
    if (phase.kind !== "done" || phase.outcome.retry) return undefined;
    const failed = phase.ran.flatMap((i, k) => (phase.outcome.results[k].ok ? [] : [{ action: checked[i].action, r: phase.outcome.results[k] }]));
    if (!failed.length) return undefined;
    return {
      error: failed.map((f) => `- ${describeAction(f.action, t, ctx)}: ${f.r.error ?? "failed"}${f.r.detail ? ` (${f.r.detail})` : ""}`).join("\n"),
      code: JSON.stringify(failed.map((f) => f.action)),
      partial: true,
    };
  })();

  /**
   * Ingen handling kan udføres (fx et script, sandkassen afviser, sammen med andre handlinger):
   * uden dette stod brugeren med en liste af grå felter og ingen vej videre. Årsagerne kan sendes
   * til modellen, som laver en ny plan.
   */
  const blocked = phase.kind === "review" && checked.length > 0 && checked.every((c) => c.problem);
  const blockedRetry = blocked
    ? {
        error: `None of the planned actions can be carried out. Make a new plan that avoids these problems:\n${checked
          .map((c) => `- ${c.action.action}: ${PROBLEM_HELP[c.problem!] ?? c.problem}`)
          .join("\n")}`,
        code: JSON.stringify(checked.map((c) => c.action)),
        blocked: true,
      }
    : undefined;

  const resultFor = (i: number) => (phase.kind === "done" ? phase.outcome.results[phase.ran.indexOf(i)] : undefined);
  const locked = phase.kind !== "review";

  return (
    <div className={s.root}>
      {plan.summary && <Body1>{plan.summary}</Body1>}
      {checked.length === 0 ? (
        <Caption1 className={s.muted}>{t("agentNoActions")}</Caption1>
      ) : (
        <>
          <Text weight="semibold">{t("agentPlanTitle")}</Text>
          {phase.kind === "review" && <Caption1 className={s.muted}>{t("agentPlanHint")}</Caption1>}
          <ul className={s.list} aria-label={t("agentPlanTitle")}>
            {checked.map((c, i) => {
              const r = resultFor(i);
              const label = describeAction(c.action, t, ctx);
              // Stedet, handlingen gælder, kan vises i dokumentet før den udføres (#22).
              const target = c.problem ? undefined : targetOf(c.action, ctx);
              // Hvad handlingen overskriver, vises, indtil den er udført (#75).
              const overwrite = phase.kind === "review" || phase.kind === "running" ? overwrites?.[i] : undefined;
              return (
                <li key={i} className={s.item}>
                  <Checkbox
                    checked={selected[i] && !c.problem}
                    disabled={!!c.problem || locked}
                    aria-label={label}
                    onChange={(_, d) => setSelected((x) => x.map((v, j) => (j === i ? !!d.checked : v)))}
                  />
                  <span className={s.itemText}>
                    <Text className={s.what}>{label}</Text>
                    {target && adapter.goTo && phase.kind !== "undone" && (
                      <Link as="button" className={s.show} onClick={() => void adapter.goTo!(target, ctx).catch(() => {})}>
                        {t("agentShowTarget")}
                      </Link>
                    )}
                    {c.action.reason && <Caption1 className={s.muted}>{c.action.reason}</Caption1>}
                    {c.action.action === "runScript" && (
                      <details className={s.code}>
                        <summary>{t("agentShowCode")}</summary>
                        <pre>{c.action.code}</pre>
                      </details>
                    )}
                    <span className={s.badges}>
                      {isDestructive(c.action) && (
                        <Badge size="small" appearance="tint" color="warning">
                          {t("agentDestructive")}
                        </Badge>
                      )}
                      {overwrite && (
                        <Badge size="small" appearance="tint" color="warning">
                          {overwriteLabel(c.action, overwrite, t)}
                        </Badge>
                      )}
                      {c.problem && (
                        <Badge size="small" appearance="tint" color="subtle">
                          {t(`agentProblem_${c.problem}` as const)}
                        </Badge>
                      )}
                    </span>
                    {r && (
                      <Caption1 className={mergeClasses(s.status, r.ok ? s.ok : s.fail)}>
                        {r.ok ? <CheckmarkCircle16Regular aria-hidden /> : <ErrorCircle16Regular aria-hidden />}
                        {r.ok
                          ? r.detail
                            ? r.detail
                            : r.count && r.count > 1
                              ? t("agentCount", { n: r.count })
                              : t("agentOk")
                          : t(`agentErr_${r.error ?? "failed"}` as const, { detail: r.detail ?? "" })}
                      </Caption1>
                    )}
                  </span>
                </li>
              );
            })}
          </ul>
        </>
      )}

      {blocked && (
        <MessageBar intent="warning">
          <MessageBarBody>{t("agentBlocked")}</MessageBarBody>
        </MessageBar>
      )}
      {checked.length > 0 && <Divider />}
      {phase.kind === "review" || phase.kind === "running" ? (
        checked.length > 0 && (
          <div className={s.bar}>
            <Button
              appearance="primary"
              size="small"
              icon={<Play16Filled />}
              disabled={count === 0 || phase.kind === "running" || counting}
              onClick={() => void execute()}
            >
              {count === 1 ? t("agentExecuteOne") : t("agentExecute", { n: count })}
            </Button>
            {blockedRetry && onRetry && (
              <Button size="small" onClick={() => onRetry(blockedRetry)}>
                {t("agentBlockedRetry")}
              </Button>
            )}
            {canTrack && <Switch label={t("agentTrack")} checked={track} disabled={phase.kind === "running"} onChange={(_, d) => setTrack(d.checked)} />}
          </div>
        )
      ) : phase.kind === "done" ? (
        <>
          <MessageBar intent={phase.outcome.results.every((r) => r.ok) ? "success" : "warning"}>
            <MessageBarBody>{t("agentDone", { ok: phase.outcome.results.filter((r) => r.ok).length, n: phase.outcome.results.length })}</MessageBarBody>
          </MessageBar>
          <div className={s.bar}>
            {phase.outcome.retry && onRetry ? (
              <Button size="small" appearance="primary" onClick={() => onRetry(phase.outcome.retry!)}>
                {t("agentRetryScript")}
              </Button>
            ) : (
              <>
                {failedRetry && onRetry && (
                  <Button size="small" appearance="primary" onClick={() => onRetry(failedRetry)}>
                    {t("agentRetryFailed")}
                  </Button>
                )}
                <Button size="small" icon={<ArrowUndo16Regular />} onClick={() => void undo()}>
                  {t("agentUndo")}
                </Button>
                {marked && (
                  <Button size="small" onClick={() => void unmark()}>
                    {t("agentUnmark")}
                  </Button>
                )}
              </>
            )}
          </div>
          {marked && <Caption1 className={s.muted}>{t("agentMarked")}</Caption1>}
          <Caption1 className={s.muted}>
            {t(
              phase.outcome.undoKind === "snapshot"
                ? "agentUndoNote"
                : phase.outcome.undoKind === "tracked"
                  ? "agentUndoTracked"
                  : phase.outcome.undoKind === "slides"
                    ? "agentUndoSlides"
                    : phase.outcome.undoKind === "text"
                      ? "agentUndoText"
                      : phase.outcome.undoKind === "script"
                        ? phase.outcome.retry
                          ? "agentScriptRolledBack"
                          : adapter.host === "word"
                            ? "agentUndoScriptWord"
                            : adapter.host === "powerpoint"
                              ? "agentUndoScriptSlides"
                              : "agentUndoScript"
                        : "agentUndoCells",
            )}
          </Caption1>
          {phase.outcome.nativeUndo && <Caption1 className={s.muted}>{t("agentUndoNative")}</Caption1>}
        </>
      ) : (
        <MessageBar intent="success">
          <MessageBarBody>{t("agentUndone")}</MessageBarBody>
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
