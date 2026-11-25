import { Button, Caption1, MessageBar, MessageBarBody, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowClockwise16Regular } from "@fluentui/react-icons";
import { MIN_AGENT_CONTEXT, OLLAMA_SMALL_DEFAULT, type ContextSource, type Endpoint, type ModelDetails } from "@miai/shared";
import { useCallback, useEffect, useState } from "react";
import { useBackend } from "../backend";
import { useI18n, type T } from "../i18n";
import { contextFixText } from "./TaskNotices";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  head: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalS },
  value: { fontWeight: tokens.fontWeightSemibold },
});

function sourceLabel(t: T, src: ContextSource | undefined): string | undefined {
  switch (src) {
    case "endpoint":
      return t("ctxSrc_endpoint");
    case "models":
      return t("ctxSrc_models");
    case "ollama-modelfile":
      return t("ctxSrc_ollamaModelfile");
    case "ollama-loaded":
      return t("ctxSrc_ollamaLoaded");
    case "lmstudio-loaded":
      return t("ctxSrc_lmstudioLoaded");
    default:
      return undefined;
  }
}

/**
 * Modellens kontekst og kapabiliteter under Indstillinger → Modeller (#73): hvor stor konteksten
 * er, hvor tallet kommer fra, og en advarsel med en konkret rettelse, når den er for lille til
 * agent og chat. Endpointen kan være ændret uden at være gemt; derfor sendes den med.
 */
export function ModelContext({ endpoint, model }: { endpoint: Endpoint; model: string }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const nf = new Intl.NumberFormat(lang === "da" ? "da-DK" : "en-GB");
  const [details, setDetails] = useState<ModelDetails | null>();
  const key = `${endpoint.baseUrl}|${endpoint.contextTokens ?? ""}|${model}`;

  const load = useCallback(() => {
    setDetails(undefined);
    api.modelDetails(endpoint, model).then(setDetails, () => setDetails(null));
    // Kun når adresse, kontekst eller model ændres – ikke ved hvert tastetryk i andre felter.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, key]);
  useEffect(load, [load]);

  if (details === null) return null;
  const caps = details
    ? [details.vision && t("capVision"), details.tools && t("capTools"), details.thinking && t("capThinking")].filter((c): c is string => !!c)
    : [];
  const low = !!details?.contextTokens && details.contextTokens < MIN_AGENT_CONTEXT;
  const ollamaUnknown = !!details && !details.contextTokens && details.server === "ollama";

  return (
    <div className={s.root} data-testid="model-context">
      <div className={s.head}>
        <Text weight="semibold">{t("ctxTitle")}</Text>
        <Button size="small" appearance="subtle" icon={<ArrowClockwise16Regular />} onClick={load} disabled={details === undefined}>
          {t("ctxRecheck")}
        </Button>
      </div>
      {details === undefined ? (
        <Spinner size="extra-tiny" label={t("ctxLoading")} labelPosition="after" />
      ) : (
        <>
          {details.contextTokens ? (
            <Caption1>
              <span className={s.value}>{t("ctxValue", { n: nf.format(details.contextTokens) })}</span>
              {sourceLabel(t, details.contextSource) ? ` – ${sourceLabel(t, details.contextSource)}.` : "."}
              {details.maxContextTokens && details.maxContextTokens > details.contextTokens
                ? ` ${t("ctxMax", { n: nf.format(details.maxContextTokens) })}`
                : ""}
            </Caption1>
          ) : (
            !ollamaUnknown && <Caption1>{t("ctxUnknown")}</Caption1>
          )}
          {caps.length > 0 && <Caption1>{t("ctxCaps", { list: caps.join(", ") })}</Caption1>}
          {details.vision === false && <Caption1>{t("capNoVision")}</Caption1>}
          {(low || ollamaUnknown) && (
            <MessageBar intent="warning">
              <MessageBarBody>
                {low ? t("ctxLow", { min: nf.format(MIN_AGENT_CONTEXT) }) : t("ctxOllamaNotLoaded", { n: nf.format(OLLAMA_SMALL_DEFAULT) })}{" "}
                {contextFixText(t, details.server)}
              </MessageBarBody>
            </MessageBar>
          )}
        </>
      )}
    </div>
  );
}
