import { Caption1, MessageBar, MessageBarBody, MessageBarTitle, makeStyles, tokens } from "@fluentui/react-components";
import { MIN_AGENT_CONTEXT, type DocContext, type ModelServer } from "@miai/shared";
import { useI18n, type T } from "../i18n";
import type { TaskState } from "../useTask";

const useStyles = makeStyles({
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  detail: { display: "block", marginTop: tokens.spacingVerticalXXS, overflowWrap: "anywhere" },
});

/**
 * Oversætter en fejl fra motoren til en forståelig overskrift (da/en) ud fra koden og HTTP-status (#3).
 * Motorens egen besked vises som detalje, fordi den indeholder fx adresse eller serverens svar.
 */
export function describeError(t: T, code?: string, status?: number): string | undefined {
  switch (code) {
    case "UPSTREAM_HTTP":
      if (status === 401 || status === 403) return t("err_auth");
      if (status === 404) return t("err_notFound");
      if (status === 413) return t("err_context");
      if (status === 429) return t("err_rateLimit");
      if (status && status >= 500) return t("err_server");
      return t("err_http");
    case "UPSTREAM_UNREACHABLE":
      return t("err_unreachable");
    case "UPSTREAM_TIMEOUT":
      return t("err_timeout");
    case "UPSTREAM_PROTOCOL":
      return t("err_protocol");
    case "NO_MODEL":
      return t("err_noModel");
    case "POLICY":
      return t("err_policy");
    case "CONTEXT_OVERFLOW":
      return t("err_context");
    case "NO_VISION":
      return t("err_noVision");
    default:
      return undefined;
  }
}

/** Hvordan konteksten gøres større i værktøjet (#73). */
export function contextFixText(t: T, server: ModelServer | undefined): string {
  return t(server === "ollama" ? "fixOllama" : server === "lmstudio" ? "fixLmStudio" : "fixGeneric");
}

/**
 * Fælles beskeder under et resultat: fejl, afkortet dokument, kontekstgrænse, afbrudt svar
 * og fremskridt ved opdeling. Bruges af både Handlinger og Chat.
 */
export function TaskNotices({ state, ctx, inputError }: { state: TaskState; ctx?: DocContext | null; inputError?: string }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const nf = new Intl.NumberFormat(lang === "da" ? "da-DK" : "en-GB");
  const title = state.error ? describeError(t, state.errorCode, state.errorStatus) : undefined;
  const done = state.status === "done";

  return (
    <div className={s.list}>
      {state.status === "running" && state.step && (
        <Caption1>{t("chunkProgress", { done: Math.min(state.step.done + 1, state.step.total), total: state.step.total })}</Caption1>
      )}
      {ctx?.truncated && (state.status === "running" || done) && (
        <MessageBar intent="warning">
          <MessageBarBody>
            {t("docTruncated", { sent: nf.format(ctx.text.replace(/\n\[… afkortet\]$/, "").length), total: nf.format(ctx.originalChars ?? ctx.text.length) })}
          </MessageBarBody>
        </MessageBar>
      )}
      {(inputError || state.error) && (
        <MessageBar intent="error">
          <MessageBarBody>
            {title && <MessageBarTitle>{title}</MessageBarTitle>}
            <span className={title ? s.detail : undefined}>{inputError ?? state.error}</span>
          </MessageBarBody>
        </MessageBar>
      )}
      {done && state.warnings?.includes("input-truncated") && (
        <MessageBar intent="warning">
          <MessageBarBody>{t("warnInputTruncated")}</MessageBarBody>
        </MessageBar>
      )}
      {done && state.warnings?.includes("output-truncated") && (
        <MessageBar intent="warning">
          <MessageBarBody>{t("warnOutputTruncated")}</MessageBarBody>
        </MessageBar>
      )}
      {done && state.warnings?.includes("thinking-ignored") && (
        <MessageBar intent="warning">
          <MessageBarBody>{t("warnThinkingIgnored")}</MessageBarBody>
        </MessageBar>
      )}
      {done && state.warnings?.includes("small-context") && state.context && (
        <MessageBar intent="warning">
          <MessageBarBody>
            {t("warnSmallContext", { n: nf.format(state.context.tokens), min: nf.format(MIN_AGENT_CONTEXT) })} {contextFixText(t, state.context.server)}
          </MessageBarBody>
        </MessageBar>
      )}
      {done && state.stopReason === "repetition" && (
        <MessageBar intent="warning">
          <MessageBarBody>{t("stoppedRepetition")}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}
