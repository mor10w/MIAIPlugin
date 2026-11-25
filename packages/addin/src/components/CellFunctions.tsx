import { Button, Caption1, Field, Input, MessageBar, MessageBarActions, MessageBarBody, makeStyles, tokens } from "@fluentui/react-components";
import { Broom20Regular, ChevronCircleDown20Regular } from "@fluentui/react-icons";
import type { Settings } from "@miai/shared";
import { useEffect, useState } from "react";
import { cellRuntime } from "../cell/functions";
import type { QueueState } from "../cell/queue";
import type { HostAdapter } from "../hosts";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  row: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center", flexWrap: "wrap" },
  narrow: { maxWidth: "8rem" },
});

/** Køens tilstand, så længe komponenten vises. Uden for Excels shared runtime er der ingen kø. */
function useQueueState(): QueueState | undefined {
  const rt = cellRuntime();
  const [state, setState] = useState<QueueState | undefined>(rt?.state);
  useEffect(() => rt?.subscribe(setState), [rt]);
  return rt ? state : undefined;
}

/**
 * Linjen øverst i task pane, mens =MIAI.ASK() arbejder (#80): hvor mange celler der er i gang og i
 * kø, og knappen, der stopper dem. Vises kun, når der er noget at stoppe.
 */
export function CellBar({ className }: { className?: string }) {
  const { t } = useI18n();
  const state = useQueueState();
  if (!state || state.running + state.queued === 0) return null;
  return (
    <MessageBar intent="info" className={className} data-testid="cell-bar">
      <MessageBarBody>{t("cellBar", { running: state.running, queued: state.queued })}</MessageBarBody>
      <MessageBarActions>
        <Button size="small" onClick={() => cellRuntime()?.stop()}>
          {t("cellStop")}
        </Button>
      </MessageBarActions>
    </MessageBar>
  );
}

/** Er regnearksfunktionerne registreret (Excel med shared runtime)? Ellers vises sektionen ikke. */
export function cellFunctionsAvailable(): boolean {
  return !!cellRuntime();
}

/** Indstillinger → Excel-funktioner: samtidige kald, cachen og modelvælgeren. */
export function CellSection({
  settings,
  setSettings,
  adapter,
  models,
  onMessage,
}: {
  settings: Settings;
  setSettings: (s: Settings) => void;
  adapter?: HostAdapter;
  models: string[];
  onMessage: (m: { intent: "success" | "error"; text: string }) => void;
}) {
  const { t } = useI18n();
  const s = useStyles();
  const rt = cellRuntime();
  const [cached, setCached] = useState(rt?.cacheSize ?? 0);
  // Tallet følger med, mens celler bliver færdige.
  useEffect(() => rt?.subscribe(() => setCached(rt.cacheSize)), [rt]);
  if (!rt) return null;

  const insertPicker = async () => {
    try {
      const address = await adapter!.insertModelPicker!(settings.defaultModel || models[0] || "");
      onMessage({ intent: "success", text: t("cellPickerDone", { address }) });
    } catch (e) {
      onMessage({ intent: "error", text: (e as Error).message });
    }
  };

  return (
    <>
      <Caption1>{t("cellIntro")}</Caption1>
      <Field label={t("cellConcurrency")} hint={t("cellConcurrencyHint")}>
        <Input
          className={s.narrow}
          type="number"
          min={1}
          max={16}
          value={String(settings.cellConcurrency)}
          onChange={(_, d) => setSettings({ ...settings, cellConcurrency: Math.min(16, Math.max(1, Math.round(Number(d.value)) || 1)) })}
        />
      </Field>
      {adapter?.insertModelPicker && (
        <Field hint={t("cellPickerHint")}>
          <div className={s.row}>
            <Button icon={<ChevronCircleDown20Regular />} onClick={() => void insertPicker()}>
              {t("cellPicker")}
            </Button>
          </div>
        </Field>
      )}
      <Field hint={t("cellCache", { n: cached })}>
        <div className={s.row}>
          <Button
            icon={<Broom20Regular />}
            disabled={cached === 0}
            onClick={() => {
              rt.clearCache();
              setCached(0);
              onMessage({ intent: "success", text: t("cellCacheCleared") });
            }}
          >
            {t("cellClearCache")}
          </Button>
        </div>
      </Field>
    </>
  );
}
