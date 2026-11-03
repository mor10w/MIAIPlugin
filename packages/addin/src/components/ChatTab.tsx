import { useEffect, useRef, useState } from "react";
import { Button, Caption1, Checkbox, MessageBar, MessageBarBody, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { Send20Regular } from "@fluentui/react-icons";
import type { ApplyMode, ChatMessage, DocContext } from "@miai/shared";
import type { HostAdapter } from "../hosts";
import { useI18n } from "../i18n";
import { TaskNotices } from "./TaskNotices";
import { extractGraphic } from "../svg";
import { useTask } from "../useTask";
import { ApplyBar } from "./ApplyBar";
import { Markdown } from "./Markdown";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", height: "100%", gap: tokens.spacingVerticalS },
  log: { flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  user: { alignSelf: "flex-end", background: tokens.colorBrandBackground2, padding: tokens.spacingHorizontalS, borderRadius: tokens.borderRadiusMedium, maxWidth: "90%", whiteSpace: "pre-wrap" },
  bot: { background: tokens.colorNeutralBackground2, padding: tokens.spacingHorizontalS, borderRadius: tokens.borderRadiusMedium },
  searching: { fontSize: tokens.fontSizeBase200, color: tokens.colorBrandForeground1, marginBottom: tokens.spacingVerticalXS },
  input: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "flex-end" },
  grow: { flex: 1 },
  footer: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  footerRight: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center" },
});

export function ChatTab({ adapter }: { adapter: HostAdapter }) {
  const { t } = useI18n();
  const s = useStyles();
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [draft, setDraft] = useState("");
  const [useContext, setUseContext] = useState(true);
  /** Sidst sendte kontekst – vises i footeren, så brugeren ser hvad modellen har set. */
  const [lastCtx, setLastCtx] = useState<DocContext | null>(null);
  const ctxRef = useRef<DocContext>({ host: adapter.host, text: "", isWholeDocument: false });
  const { state, start, stop } = useTask();
  const logRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    logRef.current?.scrollTo({ top: logRef.current.scrollHeight });
  }, [messages, state.text]);

  const send = async () => {
    const prompt = draft.trim();
    if (!prompt || state.status === "running") return;
    setDraft("");
    const history = messages;
    setMessages([...history, { role: "user", content: prompt }]);
    const ctx = useContext ? await adapter.getContext({ wholeDocumentFallback: true, overview: "stats" }).catch(() => ctxRef.current) : { host: adapter.host, text: "", isWholeDocument: false };
    ctxRef.current = ctx;
    setLastCtx(useContext ? ctx : null);
    const final = await start("chat", { context: ctx, params: { prompt }, history: history.slice(-20) });
    if (final.text) setMessages((m) => [...m, { role: "assistant", content: final.text }]);
  };

  const modes: ApplyMode[] = (["insertAfter", "replace", "comment"] as ApplyMode[]).filter((m) => adapter.supports(m));
  // "Indsæt som billede" kun når svaret faktisk indeholder en ```mermaid- eller ```svg-figur.
  const modesFor = (content: string): ApplyMode[] => (adapter.supports("image") && extractGraphic(content) ? ["image", ...modes] : modes);
  const lastAssistant = messages.map((m) => m.role).lastIndexOf("assistant");

  return (
    <div className={s.root}>
      <div className={s.log} ref={logRef}>
        {messages.map((m, i) =>
          m.role === "user" ? (
            <div key={i} className={s.user}>
              {m.content}
            </div>
          ) : (
            <div key={i} className={s.bot}>
              <Markdown text={m.content} />
              {i === lastAssistant && state.status !== "running" && <ApplyBar
                  modes={modesFor(m.content)}
                  apply={(mode) => (mode === "image" ? adapter.applyImage!(m.content, ctxRef.current) : adapter.applyText(m.content, "markdown", mode, ctxRef.current))}
                  copyText={m.content}
                />}
            </div>
          ),
        )}
        {state.status === "running" && (
          <div className={s.bot}>
            {state.searchQuery && (
              <div className={s.searching}>
                {t("searching")}: {state.searchQuery}…
              </div>
            )}
            <Markdown text={state.text || (state.searchQuery ? " " : "…")} />
          </div>
        )}
        <TaskNotices state={state} ctx={lastCtx} />
      </div>
      <div className={s.input}>
        <Textarea
          className={s.grow}
          resize="vertical"
          placeholder={t("chatPlaceholder")}
          value={draft}
          onChange={(_, d) => setDraft(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              void send();
            }
          }}
        />
        {state.status === "running" ? (
          <Button onClick={stop}>{t("stop")}</Button>
        ) : (
          <Button appearance="primary" icon={<Send20Regular />} onClick={send} aria-label={t("send")} />
        )}
      </div>
      <div className={s.footer}>
        <Checkbox label={t("useSelection")} checked={useContext} onChange={(_, d) => setUseContext(!!d.checked)} />
        <div className={s.footerRight}>
          {lastCtx?.address && useContext ? <Caption1>{t("contextSent")} {lastCtx.address}</Caption1> : null}
          <Button size="small" appearance="subtle" onClick={() => { setMessages([]); setLastCtx(null); }}>
            {t("newChat")}
          </Button>
        </div>
      </div>
      {state.model && <Caption1>{state.model}</Caption1>}
    </div>
  );
}
