import { useCallback, useEffect, useState } from "react";
import { Badge, Button, Caption1, Card, Spinner, Subtitle2, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowClockwise16Regular, Copy16Regular } from "@fluentui/react-icons";
import { corsSteps, discoveryTargets, Endpoint, type KnownTool, type ProbeResult } from "@miai/shared";
import { useBackend } from "../backend";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  head: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: tokens.spacingHorizontalS },
  item: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS, paddingTop: tokens.spacingVerticalS, borderTop: `1px solid ${tokens.colorNeutralStroke2}` },
  row: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  code: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    background: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    padding: `2px ${tokens.spacingHorizontalXS}`,
    overflowWrap: "anywhere",
  },
  grow: { flex: 1 },
});

interface Found {
  baseUrl: string;
  tools: KnownTool[];
  result?: ProbeResult;
}

/**
 * Førstegangsguide (#4): finder OpenAI-kompatible værktøjer på kendte porte, viser præcis
 * CORS-opsætning, hvis browseren ikke må læse svaret, og opretter endpointen med ét klik.
 */
export function SetupGuide({ onUse }: { onUse: (tool: KnownTool, models: string[]) => Promise<void> }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const [found, setFound] = useState<Found[]>(() => discoveryTargets());
  const [scanning, setScanning] = useState(false);
  const origin = window.location.origin;

  const probe = useCallback(
    (tool: KnownTool) => api.probe(Endpoint.parse({ id: tool.id, name: tool.name, baseUrl: tool.baseUrl, structuredOutput: tool.structuredOutput })),
    [api],
  );

  const scan = useCallback(async () => {
    setScanning(true);
    const targets = discoveryTargets();
    setFound(targets);
    const results = await Promise.all(targets.map((tg) => probe(tg.tools[0]).catch((e): ProbeResult => ({ ok: false, reason: "absent", message: (e as Error).message }))));
    setFound(targets.map((tg, i) => ({ ...tg, result: results[i] })));
    setScanning(false);
  }, [probe]);

  useEffect(() => {
    void scan();
  }, [scan]);

  const ready = found.filter((f) => f.result?.ok);
  const cors = found.filter((f) => f.result && !f.result.ok && f.result.reason === "cors");
  const other = found.filter((f) => f.result && !f.result.ok && f.result.reason !== "cors");
  const names = (f: Found) => f.tools.map((x) => x.name).join(" / ");

  return (
    <Card className={s.card}>
      <div className={s.head}>
        <Subtitle2>{t("setupTitle")}</Subtitle2>
        <Button size="small" icon={<ArrowClockwise16Regular />} disabled={scanning} onClick={() => void scan()}>
          {t("setupRescan")}
        </Button>
      </div>
      <Caption1>{t("setupIntro")}</Caption1>
      {scanning && <Spinner size="tiny" label={t("setupScanning")} />}

      {ready.map((f) => (
        <div key={f.baseUrl} className={s.item}>
          <div className={s.row}>
            <Text weight="semibold">{names(f)}</Text>
            <Badge appearance="tint" color="success">
              {t("setupReady", { n: f.result?.ok ? f.result.models.length : 0 })}
            </Badge>
          </div>
          <Caption1>{f.baseUrl}</Caption1>
          {f.tools.length > 1 && <Caption1>{t("setupSharedPort")}</Caption1>}
          <div className={s.row}>
            {f.tools.map((tool) => (
              <Button key={tool.id} size="small" appearance="primary" onClick={() => void onUse(tool, f.result?.ok ? f.result.models : [])}>
                {f.tools.length > 1 ? t("setupUseNamed", { name: tool.name }) : t("setupUse")}
              </Button>
            ))}
          </div>
        </div>
      ))}

      {cors.map((f) => (
        <div key={f.baseUrl} className={s.item}>
          <div className={s.row}>
            <Text weight="semibold">{names(f)}</Text>
            <Badge appearance="tint" color="warning">
              {t("setupNeedsCors")}
            </Badge>
          </div>
          <Caption1>{t("setupCorsIntro", { origin })}</Caption1>
          {corsSteps(f.tools[0].id, origin, lang).map((step, i) => (
            <div key={i}>
              {step.platform && <Caption1>{step.platform}: </Caption1>}
              {step.code && (
                <div className={s.code}>
                  <span className={s.grow}>{step.code}</span>
                  <Button size="small" appearance="subtle" icon={<Copy16Regular />} aria-label={t("copy")} onClick={() => navigator.clipboard?.writeText(step.code!)} />
                </div>
              )}
              <Caption1>{step.text}</Caption1>
            </div>
          ))}
        </div>
      ))}

      {!scanning && ready.length === 0 && cors.length === 0 && <Text>{t("setupNothing")}</Text>}
      {!scanning && other.length > 0 && (
        <Caption1>
          {t("setupNotFound")}: {other.map(names).join(", ")}
        </Caption1>
      )}
    </Card>
  );
}
