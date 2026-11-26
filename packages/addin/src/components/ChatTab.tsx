import { useEffect, useMemo, useRef, useState } from "react";
import { Button, Caption1, Switch, Text, Textarea, Tooltip, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import {
  ChatAdd20Regular,
  ChatSparkle20Regular,
  Flash20Regular,
  Globe16Regular,
  Send20Filled,
  Sparkle16Regular,
  Stop20Regular,
  Attach16Regular,
  Attach20Regular,
  Dismiss12Regular,
} from "@fluentui/react-icons";
import {
  attachmentsText,
  chartMaterial,
  MAX_ATTACHMENTS,
  type AgentPlan,
  type ApplyMode,
  type Attachment,
  type ChatMessage,
  type DocContext,
  type SourceRef,
} from "@miai/shared";
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
import { ACCEPT, AttachmentError, IMAGE_TYPES, toAttachment } from "../attach";
import { ui } from "../theme";
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
  attachments: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: tokens.spacingHorizontalXS },
  chip: {
    display: "inline-flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXXS,
    maxWidth: "100%",
    paddingLeft: tokens.spacingHorizontalXS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase200,
  },
  chipName: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: ui.chipNameMax },
  thumb: { width: ui.iconTileSmall, height: ui.iconTileSmall, objectFit: "cover", borderRadius: tokens.borderRadiusSmall },
  attachError: { color: tokens.colorPaletteRedForeground1, flexBasis: "100%" },
  sentFile: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXXS, fontSize: tokens.fontSizeBase200, opacity: 0.9 },
  // Markeringen hører til kontakten for kontekst: på sin egen linje lige under den, på linje med teksten.
  selection: { flexBasis: "100%", paddingLeft: tokens.spacingHorizontalS, marginTop: `calc(-1 * ${tokens.spacingVerticalXS})` },
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
  /** Filer og billeder vedhæftet beskeden (#27). Teksten går med i historikken; billeder sendes kun med selve beskeden. */
  attachments?: Attachment[];
  /** Hentet fra en tidligere samtale om dokumentet (#26): svaret var en plan, som ikke længere kan udføres herfra. */
  planned?: boolean;
};

/**
 * Rettebeskeden til et diagram, der ikke kan tegnes. Et diagram over tal (xychart) fejler af andre
 * grunde end et rutediagram (null i serierne, etiketter uden anførselstegn), så rådet er et andet.
 */
export const diagramFixKey = (answer: string) => (/^\s*xychart(?:-beta)?\b/m.test(answer) ? "diagramFixPromptChart" : "diagramFixPrompt");

/**
 * Har modellen tegnet en graf af tal (Mermaids xychart eller pie)? I Excel skal sådan en graf være
 * et rigtigt diagram i arket, så panelet tilbyder agenten, også når modellen ikke selv henviser til den.
 */
export const drawsChart = (answer: string) => /```mermaid[^\n]*\n(?:\s*%%[^\n]*\n)*\s*(?:xychart(?:-beta)?|pie)\b/.test(answer);

/**
 * Henviser svaret til agenten? Chat-prompten beder modellen om det, når brugeren vil have dokumentet
 * ændret, så det fanger de ønsker, ordgenkendelsen (looksLikeEditRequest) ikke kender.
 */
export const refersToAgent = (answer: string) => /(?<![\p{L}])agent(?:en|ens)?(?![\p{L}])/iu.test(answer);

/** Beskeden, som modellen ser den i historikken: en plan beskrives med sine handlinger og udfaldet. */
export function historyContent(m: Message): string {
  if (m.attachments?.length) {
    // Tekstfiler følger med i historikken (afkortet); billeder nævnes kun, så de ikke sendes igen og igen.
    const images = m.attachments.filter((a) => a.kind === "image").map((a) => `[image attached earlier: ${a.name}]`);
    return [m.content, ...images, attachmentsText(m.attachments).slice(0, 12_000)].filter(Boolean).join("\n\n");
  }
  // En plan fra en tidligere samtale: kun opsummeringen og udfaldet er gemt.
  if (!m.plan) return m.planned && m.log ? `${m.content}\nOutcome:\n${m.log}` : m.content;
  const actions = m.plan.actions.length ? `\nProposed actions (JSON): ${JSON.stringify(m.plan.actions).slice(0, 3_000)}` : "\n(no actions proposed)";
  const outcome = m.log ? `\nOutcome:\n${m.log}` : "\nOutcome: not carried out yet – the user has not approved the plan.";
  return `${m.content}${actions}${outcome}`.slice(0, 6_000);
}

/**
 * Chatten bliver liggende (skjult), når brugeren skifter til en anden fane, så en samtale, der er i
 * gang, ikke går tabt. `active` er, om fanen vises; hver gang den vises igen, genlæses de
 * indstillinger, chatten bruger, så ændringer fra Indstillinger slår igennem.
 */
export function ChatTab({ adapter, onUseAgent, active = true }: { adapter: HostAdapter; onUseAgent?: (req: AgentRequest) => void; active?: boolean }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const kit = useKitStyles();
  const api = useBackend();
  const [messages, setMessages] = useState<Message[]>([]);
  // Tæller op, hver gang fanen vises igen: genlæser indstillingerne nedenfor.
  const [shown, setShown] = useState(0);
  const wasActive = useRef(active);
  useEffect(() => {
    if (active && !wasActive.current) setShown((n) => n + 1);
    wasActive.current = active;
  }, [active]);
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
  }, [api, docKey, shown]);
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
  const agentOpts = useAgentOptions(`${messages.length}:${shown}`);
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
  }, [api, shown]);
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

  // En skjult fane mister sin rulleposition: den huskes og stilles tilbage, når chatten vises igen.
  // Kom der nyt indhold imens (fx svaret, der blev skrevet færdigt), rulles der til bunden.
  const scrollTop = useRef(0);
  const activeRef = useRef(active);
  activeRef.current = active;
  useEffect(() => {
    if (!activeRef.current) scrollTop.current = Number.MAX_SAFE_INTEGER;
    logRef.current?.scrollTo({ top: logRef.current.scrollHeight });
  }, [messages, state.text]);
  useEffect(() => {
    if (active && logRef.current) logRef.current.scrollTop = scrollTop.current;
  }, [active]);

  /**
   * Sender en besked. `plan` tvinger en plan fra agenten (knappen under et svar, når ordgenkendelsen
   * ikke fangede ønsket), `extra` er ekstra parametre til agenten (fejl fra et script, chatsvar som
   * indhold), og `echo: false` gentager ikke brugerens besked i samtalen.
   */
  // Vedhæftninger til den næste besked (#27): billeder og tekst fra filer.
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [attachError, setAttachError] = useState<string>();
  const fileInput = useRef<HTMLInputElement>(null);
  const attach = async (files: File[]) => {
    setAttachError(undefined);
    const added: Attachment[] = [];
    for (const file of files) {
      if (attachments.length + added.length >= MAX_ATTACHMENTS) {
        setAttachError(t("attachErr_tooMany", { n: MAX_ATTACHMENTS }));
        break;
      }
      try {
        added.push(await toAttachment(file, api.extractFile ? (f) => api.extractFile!(f) : undefined));
      } catch (err) {
        setAttachError(err instanceof AttachmentError ? t(`attachErr_${err.reason}` as const, { file: err.file }) : `${file.name}: ${(err as Error).message}`);
      }
    }
    if (added.length) setAttachments((now) => [...now, ...added].slice(0, MAX_ATTACHMENTS));
  };

  const send = async (text = draft, opts: { extra?: Record<string, string>; plan?: boolean; echo?: boolean } = {}) => {
    const prompt = text.trim();
    if (!prompt || running) return;
    const extra = opts.extra ?? {};
    setDraft("");
    const history = messages.map((m) => ({ role: m.role, content: historyContent(m) }));
    // Vedhæftninger hører til denne besked og ryddes fra skrivefeltet, når den er sendt.
    const sent = opts.echo === false ? [] : attachments;
    setAttachments([]);
    setAttachError(undefined);
    if (opts.echo !== false) setMessages([...messages, { role: "user", content: prompt, ...(sent.length ? { attachments: sent } : {}) }]);
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
        { context: ctx, params: { prompt, language: lang === "da" ? "Danish" : "English", ...extra }, history: history.slice(-12), attachments: sent },
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
    const final = await start("chat", { context: ctx, params, history: history.slice(-20), attachments: sent });
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
      <div className={s.log} ref={logRef} onScroll={(e) => active && (scrollTop.current = e.currentTarget.scrollTop)} role="log" aria-live="polite">
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
              {m.attachments?.map((a, k) => (
                <span key={k} className={s.sentFile}>
                  <Attach16Regular aria-hidden /> {a.name}
                </span>
              ))}
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
                          void send(t(r.blocked ? "agentBlockedPrompt" : r.partial ? "agentRetryFailedPrompt" : "agentRetryPrompt"), {
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
                    <Button size="small" appearance="primary" onClick={() => void send(t(diagramFixKey(m.content), { error: diagramErrors[i] }))}>
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
                  {onUseAgent &&
                    canAct &&
                    lastUserPrompt &&
                    (looksLikeEditRequest(lastUserPrompt) || refersToAgent(m.content) || (adapter.host === "excel" && drawsChart(m.content))) && (
                      <div className={s.agentSuggest}>
                        <span>
                          <Text weight="semibold">{t(adapter.host === "excel" && drawsChart(m.content) ? "chartSuggestTitle" : "agentSuggestTitle")}</Text>
                          <br />
                          <Caption1 className={s.muted}>
                            {t(adapter.host === "excel" && drawsChart(m.content) ? "chartSuggestText" : "agentSuggestText")}
                          </Caption1>
                        </span>
                        <Button
                          size="small"
                          appearance="primary"
                          icon={<Flash20Regular />}
                          // Må chatten ændre dokumentet, laves planen her i samtalen; ellers på fanen Handlinger.
                          onClick={() =>
                            act
                              ? void send(lastUserPrompt, { plan: true, echo: false, extra: { material: chartMaterial(m.content).slice(0, 3_900) } })
                              : onUseAgent({ prompt: lastUserPrompt, material: chartMaterial(m.content) })
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
        {(attachments.length > 0 || attachError) && (
          <div className={s.attachments}>
            {attachments.map((a, i) => (
              <span key={i} className={s.chip}>
                {/* Billedet er brugerens egen fil som data-URL (CSP tillader kun data: og blob:). */}
                {a.kind === "image" ? <img className={s.thumb} src={a.dataUrl} alt="" /> : <Attach16Regular aria-hidden />}
                <span className={s.chipName}>{a.name}</span>
                <Button
                  size="small"
                  appearance="transparent"
                  icon={<Dismiss12Regular />}
                  aria-label={t("attachRemove", { file: a.name })}
                  onClick={() => setAttachments((now) => now.filter((_, k) => k !== i))}
                />
              </span>
            ))}
            {attachError && (
              <Caption1 className={s.attachError} role="alert">
                {attachError}
              </Caption1>
            )}
          </div>
        )}
        <div className={s.inputRow}>
          <input
            ref={fileInput}
            type="file"
            multiple
            hidden
            accept={ACCEPT}
            onChange={(e) => {
              void attach([...(e.target.files ?? [])]);
              e.target.value = "";
            }}
          />
          <Tooltip content={t("attachHint")} relationship="description">
            <Button appearance="subtle" icon={<Attach20Regular />} aria-label={t("attach")} disabled={running} onClick={() => fileInput.current?.click()} />
          </Tooltip>
          <Textarea
            className={s.grow}
            appearance="filled-lighter"
            resize="vertical"
            placeholder={t("chatPlaceholder")}
            aria-label={t("chatPlaceholder")}
            value={draft}
            onChange={(_, d) => setDraft(d.value)}
            // Et skærmbillede kan sættes ind fra udklipsholderen.
            onPaste={(e) => {
              const files = [...e.clipboardData.files].filter((f) => IMAGE_TYPES.includes(f.type));
              if (!files.length) return;
              e.preventDefault();
              void attach(files);
            }}
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
          {useContext && selection !== undefined && (
            <Caption1 className={mergeClasses(s.muted, s.selection)} aria-live="polite">
              {selection === null
                ? t(adapter.host === "powerpoint" ? "selectionNonePpt" : "selectionNone")
                : selection.kind === "range"
                  ? t("selectionRange", { address: selection.address })
                  : t("selectionWords", { n: selection.words })}
            </Caption1>
          )}
          {knowledgeOn && <Switch className={s.switch} label={t("useKnowledge")} checked={useKnowledge} onChange={(_, d) => setUseKnowledge(d.checked)} />}
          {canAct && (
            <Tooltip content={t("chatActHint")} relationship="description">
              <Switch className={s.switch} label={t("chatAct")} checked={act} onChange={(_, d) => setAct(d.checked)} />
            </Tooltip>
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
