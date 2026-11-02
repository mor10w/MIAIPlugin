import { useMemo, useRef, useState } from "react";
import {
  Button,
  Caption1,
  Card,
  Dropdown,
  Field,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { diagramToSvg, parseTsv, stripEmphasis, tasksForHost, textFormat, type ApplyMode, type DiagramResult, type DocContext, type FormulaResult, type GrammarResult, type SlidesResult, type TaskDefinition } from "@miai/shared";
import type { HostAdapter } from "../hosts";
import { extractGraphic } from "../svg";
import { useI18n } from "../i18n";
import { useTask } from "../useTask";
import { ApplyBar } from "./ApplyBar";
import { Markdown } from "./Markdown";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  grid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: tokens.spacingHorizontalS },
  taskBtn: { justifyContent: "flex-start", textAlign: "left", height: "auto", minHeight: "36px", whiteSpace: "normal", overflowWrap: "anywhere", hyphens: "auto", minWidth: 0 },
  form: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  row: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center" },
  result: { padding: tokens.spacingHorizontalM },
  edit: { padding: `${tokens.spacingVerticalXS} 0`, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  del: { textDecoration: "line-through", color: tokens.colorPaletteRedForeground1 },
  ins: { color: tokens.colorPaletteGreenForeground1, fontWeight: tokens.fontWeightSemibold },
  tsvWrap: { overflowX: "auto" },
  tsv: {
    borderCollapse: "collapse",
    fontSize: tokens.fontSizeBase200,
    "& td": { border: `1px solid ${tokens.colorNeutralStroke2}`, padding: `2px ${tokens.spacingHorizontalXS}`, whiteSpace: "nowrap" },
  },
  diagram: { "& svg": { maxWidth: "100%", height: "auto", display: "block" } },
  mono: { fontFamily: tokens.fontFamilyMonospace, wordBreak: "break-all" },
});

export function ActionsTab({ adapter }: { adapter: HostAdapter }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const tasks = useMemo(() => tasksForHost(adapter.host), [adapter.host]);
  const [taskId, setTaskId] = useState<string>();
  const task = tasks.find((x) => x.id === taskId);
  const [params, setParams] = useState<Record<string, string>>({});
  const [inputError, setInputError] = useState<string>();
  const ctxRef = useRef<DocContext | null>(null);
  const { state, start, stop, reset } = useTask();

  const select = (tk: TaskDefinition) => {
    setTaskId(tk.id);
    setParams(Object.fromEntries(tk.params.filter((p) => p.default).map((p) => [p.name, p.default!])));
    setInputError(undefined);
    reset();
  };

  const run = async () => {
    if (!task) return;
    setInputError(undefined);
    try {
      const ctx = await adapter.getContext({ wholeDocumentFallback: task.input === "selectionOrDocument" });
      ctxRef.current = ctx;
      await start(task.id, { context: ctx, params: { language: lang === "da" ? "Danish" : "English", ...params } });
    } catch (err) {
      setInputError((err as Error).message);
    }
  };

  // Det faktiske format afhænger af konteksten (fx TSV ved oversættelse af et Excel-område).
  const format = task && ctxRef.current ? textFormat(task, ctxRef.current) : undefined;
  const supported: ApplyMode[] = task ? task.applyModes.filter((m) => adapter.supports(m)) : [];

  // Ren tekst: fjern Markdown-fremhævning, som små modeller ofte tilføjer.
  const shownText = format === "plain" ? stripEmphasis(state.text) : state.text;

  // Et TSV-svar hører hjemme i cellerne; "Indsæt som billede" kræver en figur i svaret.
  const modes: ApplyMode[] = task
    ? supported.filter((m) => (format === "tsv" ? m === "cells" || m === "comment" : m !== "cells") && (m !== "image" || extractGraphic(shownText) !== null))
    : [];

  const apply = async (mode: ApplyMode) => {
    const ctx = ctxRef.current!;
    if (!task) throw new Error("Ingen opgave");
    if (mode === "image") return adapter.applyImage!(shownText);
    if (format) return adapter.applyText(shownText, format, mode, ctx);
    if (task.id === "grammar") return adapter.applyGrammar!(state.result as GrammarResult, mode, ctx);
    if (task.id === "text-to-slides") return adapter.applySlides!(state.result as SlidesResult);
    if (task.id === "diagram") return adapter.applyDiagram!(state.result as DiagramResult, ctx);
    if (task.id === "generate-formula") return adapter.applyFormula!(state.result as FormulaResult, ctx);
    throw new Error("Ukendt resultattype");
  };

  return (
    <div className={s.root}>
      <div className={s.grid}>
        {tasks.map((tk) => (
          <Button key={tk.id} className={s.taskBtn} appearance={tk.id === taskId ? "primary" : "secondary"} onClick={() => select(tk)} title={tk.description[lang]}>
            {tk.label[lang]}
          </Button>
        ))}
      </div>

      {task && (
        <div className={s.form}>
          <Caption1>{task.description[lang]}</Caption1>
          {task.params
            .filter((p) => p.name !== "language")
            .map((p) =>
              p.kind === "select" ? (
                <Field key={p.name} label={p.label[lang]}>
                  <Dropdown
                    value={p.options?.find((o) => o.value === params[p.name])?.label[lang] ?? ""}
                    selectedOptions={params[p.name] ? [params[p.name]] : []}
                    onOptionSelect={(_, d) => setParams((x) => ({ ...x, [p.name]: d.optionValue ?? "" }))}
                  >
                    {p.options?.map((o) => (
                      <Option key={o.value} value={o.value}>
                        {o.label[lang]}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
              ) : (
                <Field key={p.name} label={p.label[lang]}>
                  <Textarea resize="vertical" rows={3} placeholder={p.placeholder?.[lang]} value={params[p.name] ?? ""} onChange={(_, d) => setParams((x) => ({ ...x, [p.name]: d.value }))} />
                </Field>
              ),
            )}
          <div className={s.row}>
            {state.status === "running" ? (
              <>
                <Button onClick={stop}>{t("stop")}</Button>
                <Spinner size="tiny" label={task.output.kind === "json" ? t("generatingStructured", { chars: state.progressChars }) : t("working")} />
              </>
            ) : (
              <Button appearance="primary" onClick={run}>
                {t("run")}
              </Button>
            )}
            {state.model && <Caption1>{state.model}</Caption1>}
          </div>
        </div>
      )}

      {(inputError || state.error) && (
        <MessageBar intent="error">
          <MessageBarBody>{inputError ?? state.error}</MessageBarBody>
        </MessageBar>
      )}

      {task && (state.text || state.result !== undefined) && (
        <Card className={s.result}>
          <ResultView task={task} format={format} text={shownText} result={state.result} />
          {state.stopReason === "repetition" && (
            <MessageBar intent="warning">
              <MessageBarBody>{t("stoppedRepetition")}</MessageBarBody>
            </MessageBar>
          )}
          {state.status === "done" && modes.length > 0 && !(task.id === "grammar" && !(state.result as GrammarResult).edits.length) && (
            <ApplyBar key={state.text + JSON.stringify(state.result)} modes={modes} apply={apply} copyText={task.output.kind === "text" ? shownText : undefined} />
          )}
        </Card>
      )}
    </div>
  );
}

function ResultView({ task, format, text, result }: { task: TaskDefinition; format?: string; text: string; result: unknown }) {
  const s = useStyles();
  const { t } = useI18n();
  if (format === "tsv") return <TsvPreview text={text} />;
  if (format) return format === "markdown" ? <Markdown text={text} /> : <Text style={{ whiteSpace: "pre-wrap" }}>{text}</Text>;
  if (!result) return null;
  if (task.id === "grammar") {
    const r = result as GrammarResult;
    if (!r.edits.length) return <Text>{t("noEdits")}</Text>;
    return (
      <div>
        {r.edits.map((e, i) => (
          <div key={i} className={s.edit}>
            <span className={s.del}>{e.original}</span> → <span className={s.ins}>{e.replacement}</span>
            <br />
            <Caption1>{e.explanation}</Caption1>
          </div>
        ))}
      </div>
    );
  }
  if (task.id === "diagram") {
    const r = result as DiagramResult;
    // diagramToSvg escaper alle labels, så indholdet er sikkert at indsætte.
    return <div className={s.diagram} dangerouslySetInnerHTML={{ __html: diagramToSvg(r) }} />;
  }
  if (task.id === "text-to-slides") {
    const r = result as SlidesResult;
    return (
      <ol>
        {r.slides.map((sl, i) => (
          <li key={i}>
            <strong>{sl.title}</strong>
            <ul>
              {sl.bullets.map((b, j) => (
                <li key={j}>{b}</li>
              ))}
            </ul>
          </li>
        ))}
      </ol>
    );
  }
  if (task.id === "generate-formula") {
    const r = result as FormulaResult;
    return (
      <div>
        <Text className={s.mono}>{r.formula}</Text>
        <Markdown text={r.explanation} />
      </div>
    );
  }
  return <pre>{JSON.stringify(result, null, 2)}</pre>;
}

/** Viser et TSV-svar som tabel, så brugeren kan se, hvordan det lander i cellerne. */
function TsvPreview({ text }: { text: string }) {
  const s = useStyles();
  const rows = parseTsv(text);
  return (
    <div className={s.tsvWrap}>
      <table className={s.tsv}>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              {r.map((c, j) => (
                <td key={j}>{c}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
