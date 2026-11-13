import { useEffect, useMemo, useRef, useState } from "react";
import {
  Body1,
  Button,
  Caption1,
  Card,
  Divider,
  Dropdown,
  Field,
  Option,
  SearchBox,
  Spinner,
  Subtitle2,
  Text,
  Textarea,
  makeStyles,
  mergeClasses,
  tokens,
  useId,
} from "@fluentui/react-components";
import { ArrowLeft20Regular, ChevronRight20Regular, Play20Filled, SearchInfo20Regular, Stop20Regular } from "@fluentui/react-icons";
import {
  diagramToSvg,
  formatForView,
  parseTsv,
  stripEmphasis,
  type AgentPlan,
  type DesignPlan,
  type ApplyMode,
  type DiagramResult,
  type DocContext,
  type FormulaResult,
  type GrammarResult,
  type SlidesResult,
  type TaskView,
} from "@miai/shared";
import type { HostAdapter } from "../hosts";
import { extractGraphic } from "../svg";
import { useBackend } from "../backend";
import { MAX_CHUNKED_CHARS } from "../hosts/types";
import { useI18n } from "../i18n";
import { CATEGORY_ORDER, taskMeta } from "../taskMeta";
import { ui, type TaskCategory } from "../theme";
import { TaskNotices } from "./TaskNotices";
import { useTask } from "../useTask";
import { useAgentOptions } from "../useAgentOptions";
import { planWithReads } from "../agentRun";
import { AgentPlanView } from "./AgentPlanView";
import { DesignPreview } from "./DesignPreview";
import type { AgentRequest } from "./ChatTab";
import { ApplyBar } from "./ApplyBar";
import { EmptyState, IconTile, SectionLabel, TextSkeleton, useKitStyles } from "./kit";
import { Markdown } from "./Markdown";

/** Søgefeltet vises, når der er flere opgaver end dette. */
const SEARCH_THRESHOLD = 8;

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  list: { display: "flex", flexDirection: "column", padding: tokens.spacingVerticalXS, gap: 0 },
  /**
   * Hele rækken er klikbar via knappens ::after-lag ("stretched link"), men knappen indeholder kun
   * titlen – så navnet svarer til den synlige etiket (WCAG 2.5.3), og beskrivelsen er en beskrivelse.
   */
  row: {
    position: "relative",
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    minWidth: 0,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusLarge,
    cursor: "pointer",
    ":hover": { backgroundColor: tokens.colorSubtleBackgroundHover },
    ":active": { backgroundColor: tokens.colorSubtleBackgroundPressed },
    ":focus-within": { outline: `${tokens.strokeWidthThick} solid ${tokens.colorStrokeFocus2}`, outlineOffset: `calc(-1 * ${tokens.strokeWidthThick})` },
  },
  rowText: { display: "flex", flexDirection: "column", flex: 1, minWidth: 0 },
  rowButton: {
    justifyContent: "flex-start",
    textAlign: "left",
    minWidth: 0,
    height: "auto",
    padding: 0,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    ":hover": { color: tokens.colorNeutralForeground1 },
    ":focus-visible": { outline: "none" },
    "::after": { content: '""', position: "absolute", inset: 0, borderRadius: tokens.borderRadiusLarge },
  },
  rowDesc: {
    color: tokens.colorNeutralForeground3,
    overflow: "hidden",
    display: "-webkit-box",
    WebkitLineClamp: "2",
    WebkitBoxOrient: "vertical",
  },
  chevron: { color: tokens.colorNeutralForeground3, flexShrink: 0 },
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, padding: tokens.spacingHorizontalM },
  taskHead: { display: "flex", gap: tokens.spacingHorizontalM, alignItems: "flex-start" },
  taskHeadText: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXXS, minWidth: 0 },
  muted: { color: tokens.colorNeutralForeground3 },
  back: { alignSelf: "flex-start" },
  runRow: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center", flexWrap: "wrap" },
  grow: { flex: 1 },
  resultHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalS },
  resultBody: {
    maxHeight: ui.resultMaxHeight,
    overflowY: "auto",
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    overflowWrap: "anywhere",
  },
  /** Agentens plan: knapperne (Udfør, Fortryd) skal altid kunne nås, så ingen indre scroll. */
  agentBody: { maxHeight: "none", overflowY: "visible" },
  plain: { whiteSpace: "pre-wrap" },
  edit: { padding: `${tokens.spacingVerticalXS} 0`, borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}` },
  del: { textDecoration: "line-through", color: tokens.colorPaletteRedForeground1 },
  ins: { color: tokens.colorPaletteGreenForeground1, fontWeight: tokens.fontWeightSemibold },
  tsvWrap: { overflowX: "auto" },
  tsv: {
    borderCollapse: "collapse",
    fontSize: tokens.fontSizeBase200,
    "& td": {
      border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
      padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalXS}`,
      whiteSpace: "nowrap",
    },
  },
  diagram: { "& svg": { maxWidth: "100%", height: "auto", display: "block" } },
  mono: { fontFamily: tokens.fontFamilyMonospace, wordBreak: "break-all" },
});

export function ActionsTab({
  adapter,
  agentRequest,
  onAgentRequestConsumed,
}: {
  adapter: HostAdapter;
  /** Chat → agent: ønsket og chatsvaret, der skal bruges som indhold. */
  agentRequest?: AgentRequest | null;
  onAgentRequestConsumed?: () => void;
}) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const kit = useKitStyles();
  const idBase = useId("action");
  const api = useBackend();
  // Det samlede katalog (indbygget + brugerdefineret, F14) filtreres på den aktive app.
  const [tasks, setTasks] = useState<TaskView[]>();
  useEffect(() => {
    // Chat har sin egen fane og vises ikke som handling.
    api.tasks().then(
      (all) => setTasks(all.filter((x) => x.hosts.includes(adapter.host) && x.id !== "chat")),
      () => setTasks([]),
    );
  }, [api, adapter.host]);
  // Chat → Agent-handover: overtag prompten fra chatten og udfyld agent-opgaven med den.
  useEffect(() => {
    if (!agentRequest || !tasks) return;
    const tk = tasks.find((x) => x.id === "agent");
    if (!tk) {
      onAgentRequestConsumed?.(); // værten har ingen agent-opgave – aflever alligevel
      return;
    }
    setTaskId(tk.id);
    const defaults = Object.fromEntries(tk.params.filter((p) => p.default).map((p) => [p.name, p.default!]));
    // Chatsvaret gives med som indhold, så agenten indsætter den tekst, brugeren har set – ikke en ny.
    setParams({ ...defaults, prompt: agentRequest.prompt, material: agentRequest.material });
    setInputError(undefined);
    reset();
    onAgentRequestConsumed?.();
  }, [agentRequest, tasks]); // eslint-disable-line react-hooks/exhaustive-deps -- engangs-handover pr. anmodning
  const [taskId, setTaskId] = useState<string>();
  const [query, setQuery] = useState("");
  const task = tasks?.find((x) => x.id === taskId);
  const [params, setParams] = useState<Record<string, string>>({});
  const [inputError, setInputError] = useState<string>();
  const ctxRef = useRef<DocContext | null>(null);
  const { state, start, stop, reset } = useTask();

  const select = (tk: TaskView) => {
    setTaskId(tk.id);
    setParams(Object.fromEntries(tk.params.filter((p) => p.default).map((p) => [p.name, p.default!])));
    setInputError(undefined);
    reset();
  };

  const back = () => {
    stop();
    setTaskId(undefined);
    setInputError(undefined);
    reset();
  };

  // Kodetilstand og eksport: slået til af brugeren og tilladt af politikken.
  const { scripts, exportSheet } = useAgentOptions(taskId);

  /** Agentens færdige plan (efter evt. læse-runder), og om den er undervejs. */
  const [agentPlan, setAgentPlan] = useState<AgentPlan>();
  const [agentBusy, setAgentBusy] = useState(false);
  const [reading, setReading] = useState<string | null>(null);

  const run = async (extra: Record<string, string> = {}) => {
    if (!task) return;
    setInputError(undefined);
    try {
      const ctx = await adapter.getContext({
        // Agenten arbejder på markeringen eller – uden markering – hele dokumentet (#6).
        wholeDocumentFallback: task.input === "selectionOrDocument" || task.id === "agent",
        // Agenten ser afsnittene med typografi, så den kan bevare formateringen.
        structure: task.id === "agent" || task.id === "design-deck",
        // Kolonnestatistik over hele arket er dyr at hente og bruges kun af dataanalysen.
        overview: task.id === "analyze-data" ? "stats" : "headers",
        // Opgaver, der deler lange tekster op, kan få meget mere end standardgrænsen.
        maxChars: task.chunking === "map-reduce" ? MAX_CHUNKED_CHARS : undefined,
      });
      ctxRef.current = ctx;
      const req = { context: ctx, params: { language: lang === "da" ? "Danish" : "English", ...params, ...extra } };
      if (task.id === "agent") {
        // Agenten kan læse i dokumentet (fx andre ark), før den laver planen (#20). Planen vises først,
        // når den er færdig – ikke mellemregningerne med læsninger.
        setAgentBusy(true);
        const final = await planWithReads(start, adapter, req, (addresses) => setReading(addresses ? addresses.join(", ") : null)).finally(() =>
          setAgentBusy(false),
        );
        setAgentPlan(final.status === "done" ? (final.result as AgentPlan | undefined) : undefined);
      } else await start(task.id, req);
    } catch (err) {
      setInputError((err as Error).message);
    }
  };

  // Det faktiske format afhænger af konteksten (fx TSV ved oversættelse af et Excel-område).
  const format = task && ctxRef.current ? formatForView(task, ctxRef.current) : undefined;
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
    if (mode === "image") return adapter.applyImage!(shownText, ctx);
    if (format) return adapter.applyText(shownText, format, mode, ctx);
    if (task.id === "grammar") return adapter.applyGrammar!(state.result as GrammarResult, mode, ctx);
    if (task.id === "text-to-slides") return adapter.applySlides!(state.result as SlidesResult);
    if (task.id === "diagram") return adapter.applyDiagram!(state.result as DiagramResult, ctx);
    if (task.id === "generate-formula") return adapter.applyFormula!(state.result as FormulaResult, ctx);
    throw new Error("Ukendt resultattype");
  };

  // --- Liste over handlinger, grupperet efter kategori ------------------------------------
  const groups = useMemo(() => {
    const q = query.trim().toLowerCase();
    const visible = (tasks ?? []).filter((tk) => !q || `${tk.label[lang]} ${tk.description[lang]}`.toLowerCase().includes(q));
    return CATEGORY_ORDER.map((cat) => ({ cat, items: visible.filter((tk) => taskMeta(tk).category === cat) })).filter((g) => g.items.length);
  }, [tasks, query, lang]);

  if (!task) {
    if (!tasks) return <Spinner size="small" label={t("loading")} />;
    const categoryName: Record<TaskCategory, string> = {
      agent: t("catAgent"),
      write: t("catWrite"),
      present: t("catPresent"),
      analyze: t("catAnalyze"),
      custom: t("catCustom"),
    };
    return (
      <div className={mergeClasses(s.root, kit.appear)}>
        {tasks.length > SEARCH_THRESHOLD && (
          <SearchBox placeholder={t("searchActions")} aria-label={t("searchActions")} value={query} onChange={(_, d) => setQuery(d.value)} />
        )}
        {groups.map(({ cat, items }) => (
          <section key={cat} aria-label={categoryName[cat]}>
            <SectionLabel>{categoryName[cat]}</SectionLabel>
            <Card className={s.list}>
              {items.map((tk) => {
                const meta = taskMeta(tk);
                const descId = `${idBase}-${tk.id}-d`;
                return (
                  <div key={tk.id} className={s.row}>
                    <IconTile icon={meta.icon} category={meta.category} />
                    <span className={s.rowText}>
                      <Button
                        appearance="transparent"
                        className={s.rowButton}
                        onClick={() => select(tk)}
                        aria-describedby={tk.description[lang] ? descId : undefined}
                      >
                        {tk.label[lang]}
                      </Button>
                      {tk.description[lang] && (
                        <Caption1 id={descId} className={s.rowDesc}>
                          {tk.description[lang]}
                        </Caption1>
                      )}
                    </span>
                    <ChevronRight20Regular className={s.chevron} aria-hidden />
                  </div>
                );
              })}
            </Card>
          </section>
        ))}
        {groups.length === 0 && <EmptyState icon={SearchInfo20Regular} title={t("noActionsFound")} text={t("noActionsFoundHint")} />}
      </div>
    );
  }

  // --- En valgt handling: parametre → kør → resultat → indsæt -----------------------------
  const meta = taskMeta(task);
  const running = state.status === "running";
  const hasResult = !!state.text || state.result !== undefined;
  const isAgent = task.id === "agent";
  const isDesign = task.id === "design-deck";
  const showApply =
    !isAgent && !isDesign && state.status === "done" && modes.length > 0 && !(task.id === "grammar" && !(state.result as GrammarResult).edits.length);

  return (
    <div className={mergeClasses(s.root, kit.appear)}>
      <Button className={s.back} appearance="subtle" size="small" icon={<ArrowLeft20Regular />} onClick={back}>
        {t("allActions")}
      </Button>

      <Card className={s.card}>
        <div className={s.taskHead}>
          <IconTile icon={meta.icon} category={meta.category} />
          <div className={s.taskHeadText}>
            <Subtitle2 as="h2">{task.label[lang]}</Subtitle2>
            {task.description[lang] && <Body1 className={s.muted}>{task.description[lang]}</Body1>}
          </div>
        </div>

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
                <Textarea
                  resize="vertical"
                  rows={3}
                  placeholder={p.placeholder?.[lang]}
                  value={params[p.name] ?? ""}
                  onChange={(_, d) => setParams((x) => ({ ...x, [p.name]: d.value }))}
                />
              </Field>
            ),
          )}

        {params.material && (
          <div className={s.runRow}>
            <Caption1 className={s.muted}>{t("agentMaterial")}</Caption1>
            <Button size="small" appearance="subtle" onClick={() => setParams(({ material: _, ...rest }) => rest)}>
              {t("agentMaterialRemove")}
            </Button>
          </div>
        )}

        <div className={s.runRow}>
          {running ? (
            <Button icon={<Stop20Regular />} onClick={stop}>
              {t("stop")}
            </Button>
          ) : (
            <Button appearance="primary" icon={<Play20Filled />} onClick={() => void run()}>
              {t("run")}
            </Button>
          )}
          <span className={s.grow} />
          {state.model && <Caption1 className={s.muted}>{state.model}</Caption1>}
        </div>
      </Card>

      <TaskNotices state={state} ctx={ctxRef.current} inputError={inputError} />

      {(hasResult || running) && !state.error && (
        <Card className={mergeClasses(s.card, kit.appear)}>
          <div className={s.resultHead}>
            <Subtitle2 as="h3">{t("result")}</Subtitle2>
            {reading && <Caption1 className={s.muted}>{t("agentReading", { what: reading })}</Caption1>}
            {running && task.output.kind === "json" && <Caption1 className={s.muted}>{t("generatingStructured", { chars: state.progressChars })}</Caption1>}
          </div>
          <div className={mergeClasses(s.resultBody, (isAgent || isDesign) && s.agentBody)} aria-live="polite" aria-busy={running}>
            {isAgent && !agentBusy && state.status === "done" && agentPlan && ctxRef.current ? (
              <AgentPlanView
                key={JSON.stringify(agentPlan)}
                plan={agentPlan}
                ctx={ctxRef.current}
                adapter={adapter}
                scripts={scripts}
                exportSheet={exportSheet}
                onRetry={(r) =>
                  void run({ previousError: r.error.slice(0, 2_000), previousCode: r.code.slice(0, 3_900), ...(r.partial ? { previousPartial: "true" } : {}) })
                }
              />
            ) : isDesign && state.status === "done" && state.result !== undefined && ctxRef.current ? (
              <DesignPreview key={JSON.stringify(state.result)} plan={state.result as DesignPlan} ctx={ctxRef.current} adapter={adapter} />
            ) : hasResult && !isAgent && !isDesign ? (
              <ResultView task={task} format={format} text={shownText} result={state.result} />
            ) : (
              <TextSkeleton label={t("working")} />
            )}
          </div>
          {showApply && (
            <>
              <Divider />
              <ApplyBar
                key={state.text + JSON.stringify(state.result)}
                modes={modes}
                apply={apply}
                copyText={task.output.kind === "text" ? shownText : undefined}
              />
            </>
          )}
        </Card>
      )}
    </div>
  );
}

function ResultView({ task, format, text, result }: { task: TaskView; format?: string; text: string; result: unknown }) {
  const s = useStyles();
  const { t } = useI18n();
  if (format === "tsv") return <TsvPreview text={text} />;
  if (format) return format === "markdown" ? <Markdown text={text} /> : <Text className={s.plain}>{text}</Text>;
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
    // eslint-disable-next-line no-restricted-syntax -- diagramToSvg escaper alle labels; SVG'en bygges af os, ikke af modellen.
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
