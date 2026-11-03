import { useMemo, useState } from "react";
import {
  Badge,
  Body1,
  Button,
  Caption1,
  Checkbox,
  Divider,
  MessageBar,
  MessageBarBody,
  Switch,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from "@fluentui/react-components";
import { ArrowUndo16Regular, CheckmarkCircle16Regular, ErrorCircle16Regular, Play16Filled } from "@fluentui/react-icons";
import { checkPlan, isDestructive, parseBlocks, type AgentAction, type AgentPlan, type DocContext } from "@miai/shared";
import type { AgentOutcome, HostAdapter } from "../hosts/types";
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
    case "setValues":
      return t("agentAct_setValues", { n: a.values.reduce((n, r) => n + r.length, 0), address: a.address });
    case "setFormula":
      return t("agentAct_setFormula", { address: a.address, formula: clip(a.formula) });
    case "clear":
      return t("agentAct_clear", { address: a.address });
  }
}

type Phase = { kind: "review" } | { kind: "running" } | { kind: "done"; outcome: AgentOutcome; ran: number[] } | { kind: "undone" };

/**
 * Agentens plan (#6): plan → bekræft → udfør. Intet ændres, før brugeren har valgt handlinger
 * og trykket Udfør. Handlinger, der ikke kan udføres (tekst findes ikke osv.), vises fravalgt
 * med årsagen. Efter udførelsen kan hele sessionen fortrydes.
 */
export function AgentPlanView({ plan, ctx, adapter }: { plan: AgentPlan; ctx: DocContext; adapter: HostAdapter }) {
  const { t } = useI18n();
  const s = useStyles();
  const checked = useMemo(() => checkPlan(plan, ctx), [plan, ctx]);
  const [selected, setSelected] = useState<boolean[]>(() => checked.map((c) => !c.problem));
  const canTrack = adapter.agentTracking?.() ?? false;
  const [track, setTrack] = useState(canTrack);
  const [phase, setPhase] = useState<Phase>({ kind: "review" });
  const [error, setError] = useState<string>();
  const count = selected.filter(Boolean).length;

  const execute = async () => {
    const ran = checked.flatMap((c, i) => (selected[i] && !c.problem ? [i] : []));
    setError(undefined);
    setPhase({ kind: "running" });
    try {
      const outcome = await adapter.applyAgent!(
        ran.map((i) => checked[i].action),
        ctx,
        { trackChanges: track },
      );
      setPhase({ kind: "done", outcome, ran });
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
      setPhase({ kind: "undone" });
    } catch (err) {
      setError((err as Error).message);
    }
  };

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
                    {c.action.reason && <Caption1 className={s.muted}>{c.action.reason}</Caption1>}
                    <span className={s.badges}>
                      {isDestructive(c.action) && (
                        <Badge size="small" appearance="tint" color="warning">
                          {t("agentDestructive")}
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
                          ? r.count && r.count > 1
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

      {checked.length > 0 && <Divider />}
      {phase.kind === "review" || phase.kind === "running" ? (
        checked.length > 0 && (
          <div className={s.bar}>
            <Button appearance="primary" size="small" icon={<Play16Filled />} disabled={count === 0 || phase.kind === "running"} onClick={() => void execute()}>
              {count === 1 ? t("agentExecuteOne") : t("agentExecute", { n: count })}
            </Button>
            {canTrack && <Switch label={t("agentTrack")} checked={track} disabled={phase.kind === "running"} onChange={(_, d) => setTrack(d.checked)} />}
          </div>
        )
      ) : phase.kind === "done" ? (
        <>
          <MessageBar intent={phase.outcome.results.every((r) => r.ok) ? "success" : "warning"}>
            <MessageBarBody>{t("agentDone", { ok: phase.outcome.results.filter((r) => r.ok).length, n: phase.outcome.results.length })}</MessageBarBody>
          </MessageBar>
          <div className={s.bar}>
            <Button size="small" icon={<ArrowUndo16Regular />} onClick={() => void undo()}>
              {t("agentUndo")}
            </Button>
          </div>
          <Caption1 className={s.muted}>
            {t(phase.outcome.undoKind === "snapshot" ? "agentUndoNote" : phase.outcome.undoKind === "tracked" ? "agentUndoTracked" : "agentUndoCells")}
          </Caption1>
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
