import { useEffect, useRef, useState } from "react";
import { Button, Caption1, Checkbox, MessageBar, MessageBarBody, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { Send20Regular } from "@fluentui/react-icons";
import type { ApplyMode, ChatMessage, DocContext } from "@miai/shared";
import type { HostAdapter } from "../hosts";
import { useI18n } from "../i18n";
import { useTask } from "../useTask";
import { ApplyBar } from "./ApplyBar";
import { Markdown } from "./Markdown";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", height: "100%", gap: tokens.spacingVerticalS },
  log: { flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  user: { alignSelf: "flex-end", background: tokens.colorBrandBackground2, padding: tokens.spacingHorizontalS, borderRadius: tokens.borderRadiusMedium, maxWidth: "90%", whiteSpace: "pre-wrap" },
  bot: { background: tokens.colorNeutralBackground2, padding: tokens.spacingHorizontalS, borderRadius: tokens.borderRadiusMedium },
  input: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "flex-end" },
  grow: { flex: 1 },
  footer: { display: "flex", justifyContent: "space-between", alignItems: "center" },
});

export function ChatTab({ adapter }: { adapter: HostAdapter }) {
  const { t } = useI18n();
  const s = useStyles();
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [draft, setDraft] = useState("");
  const [useContext, setUseContext] = useState(true);
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
    const ctx = useContext ? await adapter.getContext({ wholeDocumentFallback: true }).catch(() => ctxRef.current) : { host: adapter.host, text: "", isWholeDocument: false };
    ctxRef.current = ctx;
    const final = await start("chat", { context: ctx, params: { prompt }, history: history.slice(-20) });
    if (final.text) setMessages((m) => [...m, { role: "assistant", content: final.text }]);
  };

  const modes: ApplyMode[] = (["insertAfter", "replace", "comment"] as ApplyMode[]).filter((m) => adapter.supports(m));
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
              {i === lastAssistant && state.status !== "running" && <ApplyBar modes={modes} apply={(mode) => adapter.applyText(m.content, "markdown", mode, ctxRef.current)} copyText={m.content} />}
            </div>
          ),
        )}
        {state.status === "running" && (
          <div className={s.bot}>
            <Markdown text={state.text || "…"} />
          </div>
        )}
        {state.error && (
          <MessageBar intent="error">
            <MessageBarBody>{state.error}</MessageBarBody>
          </MessageBar>
        )}
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
        <Button size="small" appearance="subtle" onClick={() => setMessages([])}>
          {t("newChat")}
        </Button>
      </div>
      {state.model && <Caption1>{state.model}</Caption1>}
    </div>
  );
}
