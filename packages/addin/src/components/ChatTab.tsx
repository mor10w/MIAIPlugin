import { useEffect, useRef, useState } from "react";
import { Button, Caption1, Switch, Text, Textarea, Tooltip, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import { ChatAdd20Regular, ChatSparkle20Regular, Flash20Regular, Globe16Regular, Send20Filled, Sparkle16Regular, Stop20Regular } from "@fluentui/react-icons";
import type { ApplyMode, ChatMessage, DocContext, SourceRef } from "@miai/shared";
import { useBackend } from "../backend";
import type { HostAdapter } from "../hosts";
import { useI18n } from "../i18n";
import { TaskNotices } from "./TaskNotices";
import { extractGraphic } from "../svg";
import { useTask } from "../useTask";
import { ApplyBar } from "./ApplyBar";
import { EmptyState, TextSkeleton, useKitStyles } from "./kit";
import { Markdown } from "./Markdown";
import { SourceList } from "./SourceList";
import { looksLikeEditRequest } from "../editIntent";

/** Chat → agent: brugerens ønske og chatsvaret, som agenten skal bruge som indhold. */
export interface AgentRequest {
  prompt: string;
  material: string;
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", height: "100%", gap: tokens.spacingVerticalS, minHeight: 0 },
  log: {
    flex: 1,
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    paddingBottom: tokens.spacingVerticalS,
  },
  user: {
    alignSelf: "flex-end",
    maxWidth: "88%",
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: `${tokens.borderRadiusXLarge} ${tokens.borderRadiusXLarge} ${tokens.borderRadiusSmall} ${tokens.borderRadiusXLarge}`,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
  },
  bot: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    borderRadius: `${tokens.borderRadiusXLarge} ${tokens.borderRadiusXLarge} ${tokens.borderRadiusXLarge} ${tokens.borderRadiusSmall}`,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow2,
  },
  botHead: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground3 },
  searching: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, color: tokens.colorBrandForeground1 },
  suggestions: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS, width: "100%" },
  suggestion: { justifyContent: "flex-start", textAlign: "left", fontWeight: tokens.fontWeightRegular },
  composer: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusXLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
    boxShadow: tokens.shadow4,
    ":focus-within": { border: `${tokens.strokeWidthThin} solid ${tokens.colorBrandStroke1}` },
  },
  inputRow: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "flex-end" },
  grow: { flex: 1, minWidth: 0 },
  footer: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: tokens.spacingHorizontalS },
  switch: { flex: 1, minWidth: 0 },
  muted: { color: tokens.colorNeutralForeground3 },
  /** Zone med apply-knapper + agent-forslag under det seneste chat-svar. */
  applyZone: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  agentSuggest: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground2,
  },
});

export function ChatTab({ adapter, onUseAgent }: { adapter: HostAdapter; onUseAgent?: (req: AgentRequest) => void }) {
  const { t } = useI18n();
  const s = useStyles();
  const kit = useKitStyles();
  const [messages, setMessages] = useState<(ChatMessage & { model?: string; sources?: SourceRef[] })[]>([]);
  const api = useBackend();
  // Søgning i egne mapper (F15): kun når bridgen har et aktivt vidensindeks.
  const [knowledgeOn, setKnowledgeOn] = useState(false);
  const [useKnowledge, setUseKnowledge] = useState(true);
  useEffect(() => {
    api.knowledge?.status().then(
      (st) => setKnowledgeOn(st.state !== "disabled"),
      () => setKnowledgeOn(false),
    );
  }, [api]);
  const [draft, setDraft] = useState("");
  const [useContext, setUseContext] = useState(true);
  /** Sidst sendte kontekst – vises i footeren, så brugeren ser hvad modellen har set. */
  const [lastCtx, setLastCtx] = useState<DocContext | null>(null);
  const ctxRef = useRef<DocContext>({ host: adapter.host, text: "", isWholeDocument: false });
  const { state, start, stop } = useTask();
  const logRef = useRef<HTMLDivElement>(null);
  const running = state.status === "running";
  /** Seneste brugermelding – afgør, om chatten foreslår agent-tilstanden. */
  const [lastUserPrompt, setLastUserPrompt] = useState("");
  /** Førstegangs-hint om apply-knapperne: vises, indtil brugeren har trykket en. */

  useEffect(() => {
    logRef.current?.scrollTo({ top: logRef.current.scrollHeight });
  }, [messages, state.text]);

  const send = async (text = draft) => {
    const prompt = text.trim();
    if (!prompt || running) return;
    setDraft("");
    const history = messages.map(({ role, content }) => ({ role, content }));
    setMessages([...messages, { role: "user", content: prompt }]);
    setLastUserPrompt(prompt);
    const ctx = useContext
      ? await adapter.getContext({ wholeDocumentFallback: true, overview: "stats" }).catch(() => ctxRef.current)
      : { host: adapter.host, text: "", isWholeDocument: false };
    ctxRef.current = ctx;
    setLastCtx(useContext ? ctx : null);
    const params: Record<string, string> = { prompt, ...(knowledgeOn && useKnowledge ? { useKnowledge: "true" } : {}) };
    const final = await start("chat", { context: ctx, params, history: history.slice(-20) });
    if (final.text) setMessages((m) => [...m, { role: "assistant", content: final.text, model: final.model, sources: final.sources }]);
  };

  const modes: ApplyMode[] = (["insertAfter", "replace", "comment"] as ApplyMode[]).filter((m) => adapter.supports(m));
  // "Indsæt som billede" kun når svaret faktisk indeholder en ```mermaid- eller ```svg-figur.
  const modesFor = (content: string): ApplyMode[] => (adapter.supports("image") && extractGraphic(content) ? ["image", ...modes] : modes);
  const lastAssistant = messages.map((m) => m.role).lastIndexOf("assistant");

  const suggestions =
    adapter.host === "excel"
      ? [t("suggestExcel1"), t("suggestExcel2"), t("suggestExcel3")]
      : adapter.host === "powerpoint"
        ? [t("suggestPpt1"), t("suggestPpt2"), t("suggestPpt3")]
        : adapter.host === "outlook"
          ? [t("suggestMail1"), t("suggestMail2"), t("suggestMail3")]
          : [t("suggestWord1"), t("suggestWord2"), t("suggestWord3")];

  return (
    <div className={s.root}>
      <div className={s.log} ref={logRef} role="log" aria-live="polite">
        {messages.length === 0 && !running && (
          <EmptyState
            icon={ChatSparkle20Regular}
            title={t("chatEmptyTitle")}
            text={t("chatEmptyText")}
            action={
              <div className={s.suggestions}>
                {suggestions.map((q) => (
                  <Button key={q} appearance="secondary" className={s.suggestion} onClick={() => void send(q)}>
                    {q}
                  </Button>
                ))}
              </div>
            }
          />
        )}
        {messages.map((m, i) =>
          m.role === "user" ? (
            <div key={i} className={mergeClasses(s.user, kit.appear)}>
              {m.content}
            </div>
          ) : (
            <div key={i} className={mergeClasses(s.bot, kit.appear)}>
              <div className={s.botHead}>
                <Sparkle16Regular aria-hidden />
                <Caption1>{m.model ?? "MIAI"}</Caption1>
              </div>
              <Markdown text={m.content} />
              {m.sources && <SourceList sources={m.sources} api={api.knowledge} />}
              {i === lastAssistant && !running && (
                <div className={s.applyZone}>
                  <ApplyBar
                    label={t("applyLabel")}
                    modes={modesFor(m.content)}
                    apply={(mode) => {
                      return mode === "image" ? adapter.applyImage!(m.content, ctxRef.current) : adapter.applyText(m.content, "markdown", mode, ctxRef.current);
                    }}
                    copyText={m.content}
                  />
                  {onUseAgent && adapter.supports("agent") && lastUserPrompt && looksLikeEditRequest(lastUserPrompt) && (
                    <div className={s.agentSuggest}>
                      <span>
                        <Text weight="semibold">{t("agentSuggestTitle")}</Text>
                        <br />
                        <Caption1 className={s.muted}>{t("agentSuggestText")}</Caption1>
                      </span>
                      <Button
                        size="small"
                        appearance="primary"
                        icon={<Flash20Regular />}
                        onClick={() => onUseAgent({ prompt: lastUserPrompt, material: m.content })}
                      >
                        {t("agentSuggestCta")}
                      </Button>
                    </div>
                  )}
                </div>
              )}
            </div>
          ),
        )}
        {running && (
          <div className={mergeClasses(s.bot, kit.appear)}>
            <div className={s.botHead}>
              <Sparkle16Regular aria-hidden />
              <Caption1>{state.model ?? "MIAI"}</Caption1>
            </div>
            {state.searchQuery && (
              <Caption1 className={s.searching}>
                <Globe16Regular aria-hidden />
                {t("searching")}: {state.searchQuery}…
              </Caption1>
            )}
            {state.text ? <Markdown text={state.text} /> : <TextSkeleton lines={2} label={t("working")} />}
            {state.sources && <SourceList sources={state.sources} api={api.knowledge} />}
          </div>
        )}
        <TaskNotices state={state} ctx={lastCtx} />
      </div>

      <div className={s.composer}>
        <div className={s.inputRow}>
          <Textarea
            className={s.grow}
            appearance="filled-lighter"
            resize="vertical"
            placeholder={t("chatPlaceholder")}
            aria-label={t("chatPlaceholder")}
            value={draft}
            onChange={(_, d) => setDraft(d.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                void send();
              }
            }}
          />
          {running ? (
            <Button icon={<Stop20Regular />} onClick={stop} aria-label={t("stop")} />
          ) : (
            <Button appearance="primary" icon={<Send20Filled />} onClick={() => void send()} aria-label={t("send")} disabled={!draft.trim()} />
          )}
        </div>
        <div className={s.footer}>
          <Switch className={s.switch} label={t("useSelection")} checked={useContext} onChange={(_, d) => setUseContext(d.checked)} />
          {knowledgeOn && <Switch label={t("useKnowledge")} checked={useKnowledge} onChange={(_, d) => setUseKnowledge(d.checked)} />}
          {lastCtx?.address && useContext ? (
            <Caption1 className={s.muted}>
              {t("contextSent")} {lastCtx.address}
            </Caption1>
          ) : null}
          {messages.length > 0 && (
            <Tooltip content={t("newChat")} relationship="label">
              <Button
                size="small"
                appearance="subtle"
                icon={<ChatAdd20Regular />}
                onClick={() => {
                  setMessages([]);
                  setLastCtx(null);
                }}
              />
            </Tooltip>
          )}
        </div>
      </div>
    </div>
  );
}
