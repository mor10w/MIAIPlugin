import { useEffect, useMemo, useRef, useState } from "react";
import { Button, Caption1, Switch, Text, Textarea, Tooltip, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import { ChatAdd20Regular, ChatSparkle20Regular, Flash20Regular, Globe16Regular, Send20Filled, Sparkle16Regular, Stop20Regular } from "@fluentui/react-icons";
import type { AgentPlan, ApplyMode, ChatMessage, DocContext, SourceRef } from "@miai/shared";
import { useBackend } from "../backend";
import type { HostAdapter } from "../hosts";
import { documentUrl, watchSelection } from "../hosts/selection";
import type { SelectionInfo } from "../hosts/types";
import { clearChat, documentKey, loadChat, saveChat } from "../chatStore";
import { useI18n } from "../i18n";
import { TaskNotices } from "./TaskNotices";
import { extractGraphic } from "../svg";
import { useTask } from "../useTask";
import { ApplyBar } from "./ApplyBar";
import { EmptyState, TextSkeleton, useKitStyles } from "./kit";
import { Markdown } from "./Markdown";
import { SourceList } from "./SourceList";
import { looksLikeEditRequest } from "../editIntent";
import { useAgentOptions } from "../useAgentOptions";
import { planWithReads } from "../agentRun";
import { useDocRefs } from "../useDocRefs";
import { AgentPlanView } from "./AgentPlanView";

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
  // Kontakterne må ikke klemmes sammen (teksterne overlappede i et smalt panel): de ombrydes til hver sin linje.
  switch: { flexShrink: 0, maxWidth: "100%" },
  footerEnd: { marginLeft: "auto" },
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

/**
 * En besked i samtalen. Et svar kan være en plan fra agenten (#19): den vises med bekræft, udfør og
 * fortryd, og `log` (udfaldet) går med i historikken, så modellen ved, hvad der faktisk skete.
 */
type Message = ChatMessage & {
  model?: string;
  sources?: SourceRef[];
  plan?: AgentPlan;
  planCtx?: DocContext;
  log?: string;
  /** Hentet fra en tidligere samtale om dokumentet (#26): svaret var en plan, som ikke længere kan udføres herfra. */
  planned?: boolean;
};

/**
 * Henviser svaret til agenten? Chat-prompten beder modellen om det, når brugeren vil have dokumentet
 * ændret, så det fanger de ønsker, ordgenkendelsen (looksLikeEditRequest) ikke kender.
 */
export const refersToAgent = (answer: string) => /(?<![\p{L}])agent(?:en|ens)?(?![\p{L}])/iu.test(answer);

/** Beskeden, som modellen ser den i historikken: en plan beskrives med sine handlinger og udfaldet. */
export function historyContent(m: Message): string {
  // En plan fra en tidligere samtale: kun opsummeringen og udfaldet er gemt.
  if (!m.plan) return m.planned && m.log ? `${m.content}\nOutcome:\n${m.log}` : m.content;
  const actions = m.plan.actions.length ? `\nProposed actions (JSON): ${JSON.stringify(m.plan.actions).slice(0, 3_000)}` : "\n(no actions proposed)";
  const outcome = m.log ? `\nOutcome:\n${m.log}` : "\nOutcome: not carried out yet – the user has not approved the plan.";
  return `${m.content}${actions}${outcome}`.slice(0, 6_000);
}

export function ChatTab({ adapter, onUseAgent }: { adapter: HostAdapter; onUseAgent?: (req: AgentRequest) => void }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const kit = useKitStyles();
  const api = useBackend();
  const [messages, setMessages] = useState<Message[]>([]);
  // Samtalen huskes pr. dokument (#26), når det er slået til, og dokumentet har en adresse.
  const docKey = useMemo(() => documentKey(documentUrl()), []);
  const [remember, setRemember] = useState(false);
  const loaded = useRef(false);
  useEffect(() => {
    api.settings().then(
      (v) => {
        const on = v.settings.chatHistory !== false;
        setRemember(on);
        if (on && !loaded.current) {
          const stored = loadChat(docKey);
          if (stored.length) setMessages((now) => (now.length ? now : stored));
        }
        loaded.current = true;
      },
      () => (loaded.current = true),
    );
  }, [api, docKey]);
  useEffect(() => {
    if (!remember || !loaded.current) return;
    // Planer gemmes som tekst (opsummering og udfald) – ikke med dokumentets indhold.
    saveChat(
      docKey,
      messages.flatMap((m) =>
        m.role === "system" ? [] : [{ role: m.role, content: m.content, model: m.model, ...(m.plan || m.planned ? { planned: true, log: m.log } : {}) }],
      ),
    );
  }, [messages, remember, docKey]);
  // Markeringen lige nu (#26): vises ved skrivefeltet og følger brugerens markering.
  const [selection, setSelection] = useState<SelectionInfo | null | undefined>(undefined);
  useEffect(() => {
    if (!adapter.selectionInfo) return;
    let alive = true;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const refresh = () =>
      void adapter.selectionInfo!().then(
        (info) => alive && setSelection(info),
        () => alive && setSelection(undefined),
      );
    refresh();
    // Markeringen ændrer sig mange gange, mens der trækkes med musen: vent, til den står stille.
    const stop = watchSelection(() => {
      clearTimeout(timer);
      timer = setTimeout(refresh, 200);
    });
    return () => {
      alive = false;
      clearTimeout(timer);
      stop();
    };
  }, [adapter]);
  // Chatten kan selv ændre dokumentet (#19): et ændringsønske bliver til en plan i samtalen.
  const canAct = adapter.supports("agent");
  const [act, setAct] = useState(true);
  const agentOpts = useAgentOptions(messages.length);
  /** Sidst sendte kontekst – vises i footeren, så brugeren ser hvad modellen har set. */
  const [lastCtx, setLastCtx] = useState<DocContext | null>(null);
  // Henvisninger til celler og slides i svarene kan klikkes (#22).
  const docRefs = useDocRefs(adapter, lastCtx);
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
  const ctxRef = useRef<DocContext>({ host: adapter.host, text: "", isWholeDocument: false });
  const { state, start, stop } = useTask();
  const logRef = useRef<HTMLDivElement>(null);
  const running = state.status === "running";
  /** Agenten er ved at lave en plan (svaret er JSON og vises først som plan, når det er færdigt). */
  const [planning, setPlanning] = useState(false);
  /** Områder, agenten er ved at læse, før den laver planen (#20). */
  const [reading, setReading] = useState<string | null>(null);
  /** Seneste brugermelding – afgør, om chatten foreslår agent-tilstanden. */
  const [lastUserPrompt, setLastUserPrompt] = useState("");
  /** Førstegangs-hint om apply-knapperne: vises, indtil brugeren har trykket en. */

  useEffect(() => {
    logRef.current?.scrollTo({ top: logRef.current.scrollHeight });
  }, [messages, state.text]);

  /**
   * Sender en besked. `plan` tvinger en plan fra agenten (knappen under et svar, når ordgenkendelsen
   * ikke fangede ønsket), `extra` er ekstra parametre til agenten (fejl fra et script, chatsvar som
   * indhold), og `echo: false` gentager ikke brugerens besked i samtalen.
   */
  const send = async (text = draft, opts: { extra?: Record<string, string>; plan?: boolean; echo?: boolean } = {}) => {
    const prompt = text.trim();
    if (!prompt || running) return;
    const extra = opts.extra ?? {};
    setDraft("");
    const history = messages.map((m) => ({ role: m.role, content: historyContent(m) }));
    if (opts.echo !== false) setMessages([...messages, { role: "user", content: prompt }]);
    setLastUserPrompt(prompt);
    // Et ønske om at ændre dokumentet: agenten laver en plan, som vises i samtalen. Intet ændres,
    // før brugeren har godkendt den. Dokumentet hentes altid (planen skal pege på noget konkret).
    if (canAct && act && (opts.plan || looksLikeEditRequest(prompt))) {
      let ctx: DocContext;
      try {
        ctx = await adapter.getContext({ wholeDocumentFallback: true, structure: true, overview: "headers" });
      } catch (err) {
        setMessages((m) => [...m, { role: "assistant", content: (err as Error).message }]);
        return;
      }
      ctxRef.current = ctx;
      setLastCtx(ctx);
      setPlanning(true);
      // Agenten kan læse i dokumentet (fx andre ark), før den laver planen (#20).
      const final = await planWithReads(
        start,
        adapter,
        { context: ctx, params: { prompt, language: lang === "da" ? "Danish" : "English", ...extra }, history: history.slice(-12) },
        (addresses) => setReading(addresses ? addresses.join(", ") : null),
      ).finally(() => setPlanning(false));
      const plan = final.result as AgentPlan | undefined;
      if (final.status === "done" && plan) setMessages((m) => [...m, { role: "assistant", content: plan.summary, model: final.model, plan, planCtx: ctx }]);
      return;
    }
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
  // Et diagram, mermaid ikke kan tegne, kan heller ikke indsættes – så vises knappen ikke.
  const [diagramErrors, setDiagramErrors] = useState<Record<number, string>>({});
  const modesFor = (content: string, i: number): ApplyMode[] =>
    adapter.supports("image") && extractGraphic(content) && !diagramErrors[i] ? ["image", ...modes] : modes;
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
              {m.plan && m.planCtx ? (
                <AgentPlanView
                  plan={m.plan}
                  ctx={m.planCtx}
                  adapter={adapter}
                  scripts={agentOpts.scripts}
                  exportSheet={agentOpts.exportSheet}
                  onOutcome={(log) => setMessages((all) => all.map((x, k) => (k === i ? { ...x, log } : x)))}
                  onRetry={
                    i === lastAssistant
                      ? (r) =>
                          void send(t(r.partial ? "agentRetryFailedPrompt" : "agentRetryPrompt"), {
                            plan: true,
                            extra: {
                              previousError: r.error.slice(0, 2_000),
                              previousCode: r.code.slice(0, 3_900),
                              ...(r.partial ? { previousPartial: "true" } : {}),
                            },
                          })
                      : undefined
                  }
                />
              ) : (
                <Markdown
                  text={m.content}
                  refs={docRefs}
                  onDiagramError={(err) =>
                    setDiagramErrors((d) => {
                      if ((d[i] ?? null) === err) return d;
                      const next = { ...d };
                      if (err) next[i] = err;
                      else delete next[i];
                      return next;
                    })
                  }
                />
              )}
              {diagramErrors[i] && (
                <div className={s.agentSuggest}>
                  <span>
                    <Text weight="semibold">{t("diagramErrorTitle")}</Text>
                    <br />
                    <Caption1 className={s.muted}>{diagramErrors[i]}</Caption1>
                  </span>
                  {i === lastAssistant && !running && (
                    <Button size="small" appearance="primary" onClick={() => void send(t("diagramFixPrompt", { error: diagramErrors[i] }))}>
                      {t("diagramFixCta")}
                    </Button>
                  )}
                </div>
              )}
              {m.planned && !m.plan && <Caption1 className={s.muted}>{t(m.log ? "chatOldPlanDone" : "chatOldPlan")}</Caption1>}
              {m.sources && <SourceList sources={m.sources} api={api.knowledge} />}
              {i === lastAssistant && !running && !m.plan && (
                <div className={s.applyZone}>
                  <ApplyBar
                    label={t("applyLabel")}
                    modes={modesFor(m.content, i)}
                    apply={(mode) => {
                      return mode === "image" ? adapter.applyImage!(m.content, ctxRef.current) : adapter.applyText(m.content, "markdown", mode, ctxRef.current);
                    }}
                    copyText={m.content}
                  />
                  {onUseAgent && canAct && lastUserPrompt && (looksLikeEditRequest(lastUserPrompt) || refersToAgent(m.content)) && (
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
                        // Må chatten ændre dokumentet, laves planen her i samtalen; ellers på fanen Handlinger.
                        onClick={() =>
                          act
                            ? void send(lastUserPrompt, { plan: true, echo: false, extra: { material: m.content.slice(0, 3_900) } })
                            : onUseAgent({ prompt: lastUserPrompt, material: m.content })
                        }
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
            {state.text && !planning ? (
              <Markdown text={state.text} />
            ) : (
              <TextSkeleton lines={2} label={planning ? (reading ? t("agentReading", { what: reading }) : t("chatPlanning")) : t("working")} />
            )}
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
          {knowledgeOn && <Switch className={s.switch} label={t("useKnowledge")} checked={useKnowledge} onChange={(_, d) => setUseKnowledge(d.checked)} />}
          {canAct && (
            <Tooltip content={t("chatActHint")} relationship="description">
              <Switch className={s.switch} label={t("chatAct")} checked={act} onChange={(_, d) => setAct(d.checked)} />
            </Tooltip>
          )}
          {useContext && selection !== undefined && (
            <Caption1 className={s.muted} aria-live="polite">
              {selection === null
                ? t(adapter.host === "powerpoint" ? "selectionNonePpt" : "selectionNone")
                : selection.kind === "range"
                  ? t("selectionRange", { address: selection.address })
                  : t("selectionWords", { n: selection.words })}
            </Caption1>
          )}
          {messages.length > 0 && (
            <Tooltip content={t("newChat")} relationship="label">
              <Button
                className={s.footerEnd}
                size="small"
                appearance="subtle"
                icon={<ChatAdd20Regular />}
                onClick={() => {
                  setMessages([]);
                  setLastCtx(null);
                  clearChat(docKey);
                }}
              />
            </Tooltip>
          )}
        </div>
      </div>
    </div>
  );
}
