import { useState } from "react";
import { Button, MessageBar, MessageBarBody, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowUndo16Regular, Copy16Regular } from "@fluentui/react-icons";
import type { ApplyMode } from "@miai/shared";
import type { ApplyOutcome } from "../hosts/types";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  bar: { display: "flex", flexWrap: "wrap", gap: tokens.spacingHorizontalS, marginTop: tokens.spacingVerticalS },
  msg: { marginTop: tokens.spacingVerticalS },
});

interface Props {
  modes: ApplyMode[];
  apply: (mode: ApplyMode) => Promise<ApplyOutcome>;
  copyText?: string;
}

/** Knapper til at indsætte et resultat i dokumentet, med beskeder og fortryd. */
export function ApplyBar({ modes, apply, copyText }: Props) {
  const { t } = useI18n();
  const s = useStyles();
  const [busy, setBusy] = useState(false);
  const [outcome, setOutcome] = useState<ApplyOutcome & { error?: string }>();

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
    <>
      <div className={s.bar}>
        {modes.map((m, i) => (
          <Button key={m} size="small" appearance={i === 0 ? "primary" : "secondary"} disabled={busy} onClick={() => run(m)}>
            {t(`apply_${m}` as const)}
          </Button>
        ))}
        {copyText !== undefined && (
          <Button size="small" icon={<Copy16Regular />} onClick={() => navigator.clipboard?.writeText(copyText)}>
            {t("copy")}
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
    </>
  );
}
