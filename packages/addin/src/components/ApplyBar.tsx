import { useState } from "react";
import { Button, Caption1, MessageBar, MessageBarBody, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowUndo16Regular, Checkmark16Regular, Copy16Regular } from "@fluentui/react-icons";
import type { ApplyMode } from "@miai/shared";
import type { ApplyOutcome } from "../hosts/types";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  /** Tydeligt "zone"-område, så brugeren ser at her indsættes der i dokumentet. */
  zone: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
  },
  label: { fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 },
  bar: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: tokens.spacingHorizontalS },
  msg: { marginTop: tokens.spacingVerticalXS },
});

interface Props {
  modes: ApplyMode[];
  apply: (mode: ApplyMode) => Promise<ApplyOutcome>;
  copyText?: string;
  /** Vises over knapperne (fx "Indsæt i dokumentet"). */
  label?: string;
}

/** Knapper til at indsætte et resultat i dokumentet, med beskeder og fortryd. */
export function ApplyBar({ modes, apply, copyText, label }: Props) {
  const { t } = useI18n();
  const s = useStyles();
  const [busy, setBusy] = useState(false);
  const [outcome, setOutcome] = useState<ApplyOutcome & { error?: string }>();
  const [copied, setCopied] = useState(false);

  const run = async (mode: ApplyMode) => {
    setBusy(true);
    try {
      const o = await apply(mode);
      setOutcome({ ...o, message: o.message ?? t("applied") });
    } catch (err) {
      setOutcome({ error: (err as Error).message });
    } finally {
      setBusy(false);
    }
  };

  const undo = async () => {
    if (!outcome?.undo) return;
    setBusy(true);
    try {
      await outcome.undo();
      setOutcome({ message: t("undone") });
    } catch (err) {
      setOutcome({ error: (err as Error).message });
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={s.zone}>
      {label && <Caption1 className={s.label}>{label}</Caption1>}
      <div className={s.bar}>
        {modes.map((m, i) => (
          <Button key={m} size="small" appearance={i === 0 ? "primary" : "secondary"} disabled={busy} onClick={() => run(m)}>
            {t(`apply_${m}` as const)}
          </Button>
        ))}
        {copyText !== undefined && (
          <Button
            size="small"
            appearance="subtle"
            icon={copied ? <Checkmark16Regular /> : <Copy16Regular />}
            onClick={() => {
              void navigator.clipboard?.writeText(copyText).then(() => {
                setCopied(true);
                setTimeout(() => setCopied(false), 1_500);
              });
            }}
          >
            {copied ? t("copied") : t("copy")}
          </Button>
        )}
        {outcome?.undo && (
          <Button size="small" icon={<ArrowUndo16Regular />} disabled={busy} onClick={undo}>
            {t("undo")}
          </Button>
        )}
      </div>
      {outcome && (outcome.message || outcome.error) && (
        <MessageBar className={s.msg} intent={outcome.error ? "error" : outcome.warning ? "warning" : "success"}>
          <MessageBarBody>{outcome.error ?? outcome.message}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}
