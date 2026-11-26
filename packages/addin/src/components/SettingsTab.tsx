import { useEffect, useState } from "react";
import { policyReloadPending, syncCspAfterSave } from "../csp";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Button,
  Caption1,
  Card,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Option,
  Radio,
  Spinner,
  Switch,
  Text,
  Textarea,
  Tooltip,
  dropdownClassNames,
  inputClassNames,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  KNOWN_TOOLS,
  type AuditEntry,
  type Endpoint,
  type KnownTool,
  type Policy,
  type ProbeResult,
  type ReasoningEffort,
  type Settings,
  type StructuredOutputMode,
  type ThinkingMode,
} from "@miai/shared";
import {
  Add20Regular,
  BrainCircuit20Regular,
  Delete20Regular,
  Globe20Regular,
  History20Regular,
  Notepad20Regular,
  PlugConnected20Regular,
  Save20Regular,
  Search20Regular,
  Shield20Regular,
  TextEditStyle20Regular,
  BookSearch20Regular,
  TableCalculator20Regular,
} from "@fluentui/react-icons";
import { useBackend } from "../backend";
import { ui } from "../theme";
import { useI18n } from "../i18n";
import { SetupGuide } from "./SetupGuide";
import { CellSection, cellFunctionsAvailable } from "./CellFunctions";
import type { HostAdapter } from "../hosts";
import { ModelContext } from "./ModelContext";
import { TaskModels } from "./TaskModels";
import { KnowledgeSection } from "./KnowledgeSection";
import { InstructionsSection } from "./InstructionsSection";

/** Kendte værktøjer (fælles liste i shared, #4) + "Andet" til alt andet. */
const PRESETS: KnownTool[] = [...KNOWN_TOOLS, { id: "custom", name: "Andet / Other", baseUrl: "http://127.0.0.1:8000/v1", structuredOutput: "none" }];

/** presence_penalty-forudindstillinger. "" = sendes ikke (serverens standard). */
const THINKING_MODES: ThinkingMode[] = ["default", "on", "off"];
const EFFORTS: ReasoningEffort[] = ["default", "low", "medium", "high"];

const PENALTY_PRESETS = [
  { value: "", label: "penaltyOff" },
  { value: "0.5", label: "penaltyLight" },
  { value: "1", label: "penaltyMedium" },
  { value: "1.5", label: "penaltyStrong" },
] as const;

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, minWidth: 0 },
  accordion: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  section: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusLarge,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
    overflow: "hidden",
  },
  panel: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, paddingBottom: tokens.spacingVerticalM },
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  endpointHead: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, minWidth: 0 },
  row: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center", flexWrap: "wrap" },
  twoCol: {
    display: "grid",
    gridTemplateColumns: `repeat(auto-fit, minmax(${ui.colMin}, 1fr))`,
    gap: tokens.spacingHorizontalM,
    // Ellers strækkes et Field uden hint til naboens højde, og Fields egne rækker deler den ekstra plads, så feltet glider ned.
    alignItems: "start",
    // Input (size=20, ca. 240 px) og Dropdown (min-width 250 px) er ellers bredere end kolonnen og løber ud over kortet.
    [`& .${inputClassNames.root}`]: { minWidth: 0 },
    [`& .${dropdownClassNames.root}`]: { minWidth: 0 },
  },
  grow: { flex: 1 },
  muted: { color: tokens.colorNeutralForeground3 },
  logCard: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  logWrap: { overflowX: "auto", maxHeight: ui.resultMaxHeight, overflowY: "auto" },
  logTable: {
    borderCollapse: "collapse",
    fontSize: tokens.fontSizeBase200,
    width: "100%",
    "& td, & th": {
      border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
      padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalXS}`,
      whiteSpace: "nowrap",
      textAlign: "left",
    },
  },
  saveBar: {
    position: "sticky",
    bottom: 0,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: `${tokens.spacingVerticalS} 0`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
});

function uniqueId(base: string, taken: string[]): string {
  let id = base;
  for (let i = 2; taken.includes(id); i++) id = `${base}-${i}`;
  return id;
}

export function SettingsTab({
  onLanguage,
  onSaved,
  adapter,
}: {
  onLanguage: (l: "da" | "en") => void;
  onSaved?: () => void;
  /** Værten, til det, der kun findes i ét program (fx modelvælgeren i Excel). */
  adapter?: HostAdapter;
}) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const [settings, setSettings] = useState<Settings>();
  const [policy, setPolicy] = useState<Policy>();
  const [policySource, setPolicySource] = useState<string>("default");
  /** Tæller gem, så vidensindeksets status hentes igen (indekseringen starter ved gem). */
  const [saves, setSaves] = useState(0);
  const [hostsText, setHostsText] = useState("");
  const [models, setModels] = useState<string[]>([]);
  const [msg, setMsg] = useState<{ intent: "success" | "error" | "info"; text: string }>();
  // Task pane er genindlæst, fordi politikkens hosts blev ændret (#28): sig det, så brugeren kan teste forbindelsen igen.
  useEffect(() => {
    if (policyReloadPending(true)) setMsg({ intent: "success", text: t("policyReloaded") });
    // Kun ved første visning.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const [testing, setTesting] = useState(false);
  const [preset, setPreset] = useState(PRESETS[1].id);
  const [showAudit, setShowAudit] = useState(false);
  /** Guiden vises automatisk, når den aktive endpoint ikke svarer (#4). */
  const [showGuide, setShowGuide] = useState(false);
  /** Gem ventede på bekræftelse, fordi endpointen ikke svarede. */
  const [unverified, setUnverified] = useState<ProbeResult | null>(null);
  const [auditEntries, setAuditEntries] = useState<AuditEntry[]>();

  /** Henter loggen, når visningen åbnes. Posterne vises nyeste først. */
  const toggleAudit = async () => {
    if (showAudit) return setShowAudit(false);
    try {
      const entries = await api.audit();
      setAuditEntries([...entries].reverse());
      setShowAudit(true);
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    }
  };

  useEffect(() => {
    api
      .settings()
      .then((v) => {
        setSettings(v.settings);
        setPolicy(v.policy);
        setPolicySource(v.policySource);
        setHostsText(v.policy.allowedHosts.join("\n"));
      })
      .catch((e) => setMsg({ intent: "error", text: (e as Error).message }));
    api.models().then(
      (m) => setModels(m.map((x) => x.id)),
      () => {
        setModels([]);
        setShowGuide(true);
      },
    );
  }, [api]);

  if (!settings || !policy)
    return msg ? (
      <MessageBar intent="error">
        <MessageBarBody>{msg.text}</MessageBarBody>
      </MessageBar>
    ) : (
      <Spinner />
    );

  const activeEp = settings.endpoints.find((e) => e.id === settings.activeEndpoint);
  /** Viser også en værdi, der er sat manuelt i config.json og ikke er en af forudindstillingerne. */
  const thinkingLabel = (m: ThinkingMode) => t(m === "on" ? "thinkingOn" : m === "off" ? "thinkingOff" : "thinkingDefault");
  const effortLabel = (e: ReasoningEffort) => t(e === "low" ? "effortLow" : e === "medium" ? "effortMedium" : e === "high" ? "effortHigh" : "thinkingDefault");
  const penaltyLabel = (v: number | undefined) => {
    const preset = PENALTY_PRESETS.find((p) => (p.value ? Number(p.value) : undefined) === v);
    return preset ? t(preset.label) : String(v);
  };
  const locked = policy.locked || policySource === "machine" || policySource === "organization";
  const policySourceLabel =
    {
      default: t("policySrcDefault"),
      user: t("policySrcUser"),
      machine: t("policySrcMachine"),
      browser: t("policySrcBrowser"),
      organization: t("policySrcOrganization"),
    }[policySource] ?? policySource;
  const setEp = (id: string, patch: Partial<Endpoint>) =>
    setSettings({ ...settings, endpoints: settings.endpoints.map((e) => (e.id === id ? { ...e, ...patch } : e)) });

  const addEndpoint = () => {
    const p = PRESETS.find((x) => x.id === preset)!;
    const id = uniqueId(
      p.id,
      settings.endpoints.map((e) => e.id),
    );
    const ep: Endpoint = {
      id,
      name: p.name,
      baseUrl: p.baseUrl,
      apiKey: "",
      structuredOutput: p.structuredOutput,
      stripThinking: true,
      extraBody: {},
      timeoutMs: 300_000,
    };
    setSettings({ ...settings, endpoints: [...settings.endpoints, ep] });
  };

  /** Gemmer – men tester først den aktive endpoint, så en forkert adresse ikke gemmes ubemærket (#4). */
  const save = async (force = false, next: Settings = settings) => {
    setMsg(undefined);
    setUnverified(null);
    const nextPolicy = locked
      ? undefined
      : {
          ...policy,
          allowedHosts: hostsText
            .split(/\s+/)
            .map((h) => h.trim())
            .filter(Boolean),
        };
    if (!force) {
      const active = next.endpoints.find((e) => e.id === next.activeEndpoint);
      const probe = active ? await api.probe(active).catch((e): ProbeResult => ({ ok: false, reason: "absent", message: (e as Error).message })) : null;
      if (probe && !probe.ok && probe.reason !== "policy") {
        setUnverified(probe);
        return;
      }
      if (probe?.ok) setModels(probe.models);
    }
    try {
      await api.saveSettings(next, nextPolicy);
      // Nye hosts i politikken kræver en genindlæsning, før browseren tillader dem (CSP, #28).
      if (nextPolicy && syncCspAfterSave(api.mode, nextPolicy)) return;
      if (nextPolicy) setPolicy(nextPolicy);
      onLanguage(next.uiLanguage);
      onSaved?.();
      setSaves((n) => n + 1);
      setMsg({ intent: "success", text: t("saved") });
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    }
  };

  /** Fra guiden: brug et fundet værktøj – eksisterende endpoint med samme adresse genbruges. */
  const useTool = async (tool: KnownTool, found: string[]) => {
    const existing = settings.endpoints.find((e) => e.baseUrl.replace(/\/+$/, "") === tool.baseUrl);
    const ep: Endpoint = existing ?? {
      id: uniqueId(
        tool.id,
        settings.endpoints.map((e) => e.id),
      ),
      name: tool.name,
      baseUrl: tool.baseUrl,
      apiKey: "",
      structuredOutput: tool.structuredOutput,
      stripThinking: true,
      extraBody: {},
      timeoutMs: 300_000,
    };
    const next: Settings = { ...settings, endpoints: existing ? settings.endpoints : [...settings.endpoints, ep], activeEndpoint: ep.id, defaultModel: "" };
    setSettings(next);
    setModels(found);
    await save(true, next);
    setShowGuide(false);
    setMsg({ intent: "success", text: t("setupConnected", { name: tool.name, n: found.length }) });
  };

  const test = async () => {
    setTesting(true);
    setMsg(undefined);
    try {
      // Gem først, så bridgen tester den aktuelle konfiguration.
      const nextPolicy = locked ? undefined : { ...policy, allowedHosts: hostsText.split(/\s+/).filter(Boolean) };
      await api.saveSettings(settings, nextPolicy);
      if (nextPolicy && syncCspAfterSave(api.mode, nextPolicy)) return;
      const m = await api.models();
      setModels(m.map((x) => x.id));
      setMsg({ intent: "success", text: t("modelsFound", { n: m.length }) });
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    } finally {
      setTesting(false);
    }
  };

  const endpointAdvanced = (ep: Endpoint) => (
    <>
      <Field label={t("structuredOutput")} hint={t("structuredOutputHint")}>
        <Dropdown
          value={ep.structuredOutput}
          selectedOptions={[ep.structuredOutput]}
          onOptionSelect={(_, d) => setEp(ep.id, { structuredOutput: d.optionValue as StructuredOutputMode })}
        >
          <Option value="json_schema">json_schema</Option>
          <Option value="json_object">json_object</Option>
          <Option value="none">none</Option>
        </Dropdown>
      </Field>
      <Field label={t("presencePenalty")} hint={t("presencePenaltyHint")}>
        <Dropdown
          value={penaltyLabel(ep.presencePenalty)}
          selectedOptions={[String(ep.presencePenalty ?? "")]}
          onOptionSelect={(_, d) => setEp(ep.id, { presencePenalty: d.optionValue ? Number(d.optionValue) : undefined })}
        >
          {PENALTY_PRESETS.map((p) => (
            <Option key={p.value} value={p.value} text={penaltyLabel(p.value ? Number(p.value) : undefined)}>
              {penaltyLabel(p.value ? Number(p.value) : undefined)}
            </Option>
          ))}
        </Dropdown>
      </Field>
      <Field label={t("contextTokens")} hint={t("contextTokensHint")}>
        <Input
          type="number"
          min={512}
          step={1024}
          placeholder="auto"
          value={ep.contextTokens ? String(ep.contextTokens) : ""}
          onChange={(_, d) => setEp(ep.id, { contextTokens: d.value ? Math.max(512, Math.round(Number(d.value))) || undefined : undefined })}
        />
      </Field>
      {/* Tid uden svar (før svaret begynder og mellem stykkerne) – et svar, der strømmer, afbrydes ikke. */}
      <Field label={t("timeout")} hint={t("timeoutHint")}>
        <TimeoutInput valueMs={ep.timeoutMs} onCommit={(ms) => setEp(ep.id, { timeoutMs: ms })} suffix={t("seconds")} />
      </Field>
      <Switch label={t("stripThinking")} checked={ep.stripThinking} onChange={(_, d) => setEp(ep.id, { stripThinking: d.checked })} />
    </>
  );

  return (
    <div className={s.root}>
      <Accordion multiple collapsible defaultOpenItems={["connection"]} className={s.accordion}>
        {/* --- Forbindelse ------------------------------------------------------------- */}
        <AccordionItem value="connection" className={s.section}>
          <AccordionHeader icon={<PlugConnected20Regular />} size="large">
            {t("secConnection")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            {/* CORS-vejledningen står præcist pr. værktøj i guiden nedenfor (#4). */}
            <Caption1 className={s.muted}>{api.mode === "bridge" ? t("modeBridgeInfo") : t("modeDirectInfo")}</Caption1>
            {showGuide ? (
              <SetupGuide onUse={useTool} />
            ) : (
              <div>
                <Button size="small" icon={<Search20Regular />} onClick={() => setShowGuide(true)}>
                  {t("setupOpen")}
                </Button>
              </div>
            )}
            {settings.endpoints.map((ep) => (
              <Card key={ep.id} className={s.card} appearance="outline">
                <div className={s.endpointHead}>
                  <Radio
                    label={<Text weight="semibold">{ep.name || t("name")}</Text>}
                    aria-label={`${t("active")}: ${ep.name}`}
                    checked={settings.activeEndpoint === ep.id}
                    onChange={() => setSettings({ ...settings, activeEndpoint: ep.id, defaultModel: "" })}
                  />
                  {settings.activeEndpoint === ep.id && (
                    <Badge appearance="tint" color="brand" size="small">
                      {t("active")}
                    </Badge>
                  )}
                  <span className={s.grow} />
                  {settings.endpoints.length > 1 && (
                    <Tooltip content={t("remove")} relationship="label">
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<Delete20Regular />}
                        onClick={() => {
                          const rest = settings.endpoints.filter((e) => e.id !== ep.id);
                          setSettings({
                            ...settings,
                            endpoints: rest,
                            activeEndpoint: settings.activeEndpoint === ep.id ? rest[0].id : settings.activeEndpoint,
                          });
                        }}
                      />
                    </Tooltip>
                  )}
                </div>
                <Field label={t("name")}>
                  <Input value={ep.name} onChange={(_, d) => setEp(ep.id, { name: d.value })} />
                </Field>
                <Field label={t("baseUrl")}>
                  <Input value={ep.baseUrl} onChange={(_, d) => setEp(ep.id, { baseUrl: d.value })} />
                </Field>
                <Field label={t("apiKey")}>
                  <Input type="password" value={ep.apiKey} onChange={(_, d) => setEp(ep.id, { apiKey: d.value })} />
                </Field>
                <Accordion collapsible>
                  <AccordionItem value="advanced">
                    <AccordionHeader size="small" expandIconPosition="start">
                      {t("endpointAdvanced")}
                    </AccordionHeader>
                    <AccordionPanel className={s.panel}>{endpointAdvanced(ep)}</AccordionPanel>
                  </AccordionItem>
                </Accordion>
              </Card>
            ))}
            <div className={s.row}>
              <Dropdown
                aria-label={t("preset")}
                value={PRESETS.find((p) => p.id === preset)?.name}
                selectedOptions={[preset]}
                onOptionSelect={(_, d) => setPreset(d.optionValue ?? "custom")}
              >
                {PRESETS.map((p) => (
                  <Option key={p.id} value={p.id}>
                    {p.name}
                  </Option>
                ))}
              </Dropdown>
              <Button icon={<Add20Regular />} onClick={addEndpoint}>
                {t("addEndpoint")}
              </Button>
            </div>
            <div className={s.row}>
              <Button icon={<PlugConnected20Regular />} onClick={() => void test()} disabled={testing}>
                {t("testConnection")}
              </Button>
              {testing && <Spinner size="tiny" />}
            </div>
          </AccordionPanel>
        </AccordionItem>

        {/* --- Modeller ---------------------------------------------------------------- */}
        <AccordionItem value="models" className={s.section}>
          <AccordionHeader icon={<BrainCircuit20Regular />} size="large">
            {t("secModels")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            <Field label={t("defaultModel")}>
              <Dropdown
                value={settings.defaultModel || t("firstAvailable")}
                selectedOptions={[settings.defaultModel]}
                onOptionSelect={(_, d) => setSettings({ ...settings, defaultModel: d.optionValue ?? "" })}
              >
                <Option value="">{t("firstAvailable")}</Option>
                {models.map((m) => (
                  <Option key={m} value={m}>
                    {m}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            {activeEp && (settings.defaultModel || models[0]) && <ModelContext endpoint={activeEp} model={settings.defaultModel || models[0]} />}
            <TaskModels groups={settings.groupModels} tasks={settings.taskModels} models={models} onChange={(next) => setSettings({ ...settings, ...next })} />
          </AccordionPanel>
        </AccordionItem>

        {/* --- Svar -------------------------------------------------------------------- */}
        <AccordionItem value="replies" className={s.section}>
          <AccordionHeader icon={<TextEditStyle20Regular />} size="large">
            {t("secReplies")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            <div className={s.twoCol}>
              <Field label={t("temperature")} hint={t("temperatureHint")}>
                <Input
                  type="number"
                  min={0}
                  max={2}
                  step={0.1}
                  value={String(settings.temperature)}
                  onChange={(_, d) => setSettings({ ...settings, temperature: Number(d.value) })}
                />
              </Field>
              <Field label={t("maxTokens")}>
                <Input
                  type="number"
                  min={16}
                  step={256}
                  value={String(settings.maxTokens)}
                  onChange={(_, d) => setSettings({ ...settings, maxTokens: Number(d.value) })}
                />
              </Field>
            </div>
            {/* Tænkning og indsats (ADR-0003, tillæg): sendes i flere dialekter; serverens standard sender intet.
                Under hinanden: to rullemenuer er for brede til to kolonner i task pane. */}
            <Field label={t("thinking")} hint={t("thinkingHint")}>
              <Dropdown
                value={thinkingLabel(settings.thinking)}
                selectedOptions={[settings.thinking]}
                onOptionSelect={(_, d) => setSettings({ ...settings, thinking: (d.optionValue as ThinkingMode) ?? "default" })}
              >
                {THINKING_MODES.map((m) => (
                  <Option key={m} value={m}>
                    {thinkingLabel(m)}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label={t("reasoningEffort")} hint={t("reasoningEffortHint")}>
              <Dropdown
                disabled={settings.thinking === "off"}
                value={effortLabel(settings.reasoningEffort)}
                selectedOptions={[settings.reasoningEffort]}
                onOptionSelect={(_, d) => setSettings({ ...settings, reasoningEffort: (d.optionValue as ReasoningEffort) ?? "default" })}
              >
                {EFFORTS.map((e) => (
                  <Option key={e} value={e}>
                    {effortLabel(e)}
                  </Option>
                ))}
              </Dropdown>
            </Field>
          </AccordionPanel>
        </AccordionItem>

        {/* --- Privatliv og politik ---------------------------------------------------- */}
        <AccordionItem value="privacy" className={s.section}>
          <AccordionHeader icon={<Shield20Regular />} size="large">
            {t("secPrivacy")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            {/* Samtalen huskes pr. dokument (#26) i task panes lager på maskinen. */}
            <Field hint={t("chatHistoryHint")}>
              <Switch
                label={t("chatHistory")}
                checked={settings.chatHistory !== false}
                onChange={(_, d) => setSettings({ ...settings, chatHistory: d.checked })}
              />
            </Field>
            {/* Kodetilstand (ADR-0017): fra som standard, kan forbydes af politikken. */}
            <Field hint={policy.agentScripts === false ? t("agentScriptsForbidden") : t("agentScriptsHint")}>
              <Switch
                label={t("agentScripts")}
                disabled={policy.agentScripts === false}
                checked={settings.agentScripts && policy.agentScripts !== false}
                onChange={(_, d) => setSettings({ ...settings, agentScripts: d.checked })}
              />
            </Field>
            {/* Eksport af et ark som .xlsx (tillæg ADR-0012): kræver bridgen og kan forbydes af politikken. */}
            {api.mode === "bridge" && (
              <Field hint={policy.agentExport === false ? t("agentExportForbidden") : t("agentExportHint")}>
                <Switch
                  label={t("agentExport")}
                  disabled={policy.agentExport === false}
                  checked={settings.agentExport && policy.agentExport !== false}
                  onChange={(_, d) => setSettings({ ...settings, agentExport: d.checked })}
                />
              </Field>
            )}
            {/* Web-søgning kræver bridgen (direkte tilstand har ingen søge-eksekverer) og kan forbydes af politikken. */}
            {api.mode === "bridge" && (
              <Field hint={policy.webSearch ? t("webSearchHint") : t("webSearchForbidden")}>
                <Switch
                  label={t("webSearch")}
                  disabled={!policy.webSearch}
                  checked={settings.webSearch && policy.webSearch}
                  onChange={(_, d) => setSettings({ ...settings, webSearch: d.checked })}
                />
              </Field>
            )}
            {/* Søgetjeneste (#12): DuckDuckGo, selvhostet SearXNG eller Exa; kan låses af organisationens politik. */}
            {api.mode === "bridge" && policy.webSearch && settings.webSearch && (
              <div className={s.twoCol}>
                <Field
                  label={t("searchProvider")}
                  hint={policy.search ? t("searchLocked") : (policy.search ?? settings.search).provider === "exa" ? t("exaHint") : undefined}
                >
                  <Dropdown
                    disabled={!!policy.search}
                    value={{ duckduckgo: "DuckDuckGo", searxng: "SearXNG", exa: "Exa" }[(policy.search ?? settings.search).provider]}
                    selectedOptions={[(policy.search ?? settings.search).provider]}
                    onOptionSelect={(_, d) =>
                      setSettings({ ...settings, search: { ...settings.search, provider: d.optionValue as "duckduckgo" | "searxng" | "exa" } })
                    }
                  >
                    <Option value="duckduckgo">DuckDuckGo</Option>
                    <Option value="searxng">SearXNG</Option>
                    <Option value="exa">Exa</Option>
                  </Dropdown>
                </Field>
                {(policy.search ?? settings.search).provider === "searxng" && (
                  <Field label={t("searxngUrl")} hint={policy.search ? undefined : t("searxngHint")}>
                    <Input
                      disabled={!!policy.search}
                      placeholder="https://search.firma.local"
                      value={(policy.search ?? settings.search).url ?? ""}
                      onChange={(_, d) => setSettings({ ...settings, search: { ...settings.search, url: d.value || undefined } })}
                    />
                  </Field>
                )}
              </div>
            )}
            <Switch label={t("auditLog")} checked={settings.auditLog} onChange={(_, d) => setSettings({ ...settings, auditLog: d.checked })} />
            <div>
              <Button size="small" icon={<History20Regular />} onClick={() => void toggleAudit()}>
                {showAudit ? t("hideAudit") : t("viewAudit")}
              </Button>
            </div>
            {showAudit && (
              <div className={s.logCard}>
                {!auditEntries ? (
                  <Spinner size="tiny" />
                ) : auditEntries.length === 0 ? (
                  <Caption1 className={s.muted}>{t("auditEmpty")}</Caption1>
                ) : (
                  <div className={s.logWrap}>
                    <table className={s.logTable}>
                      <thead>
                        <tr>
                          <th>{t("auditTime")}</th>
                          <th>{t("auditTask")}</th>
                          <th>host</th>
                          <th>{t("model")}</th>
                          <th>{t("auditStatus")}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {auditEntries.map((e, i) => (
                          <tr key={i}>
                            <td>{new Date(e.ts).toLocaleString(lang === "da" ? "da-DK" : "en-GB")}</td>
                            <td>{e.kind === "task" ? e.task : e.kind}</td>
                            <td>{e.host}</td>
                            <td>{e.model ?? "—"}</td>
                            <td>
                              {e.status === "ok"
                                ? `${t("statusOk")}${e.durationMs !== undefined ? ` · ${(Math.round(e.durationMs / 10) / 100).toLocaleString(lang === "da" ? "da-DK" : "en-GB")}s` : ""}`
                                : `${t("statusError")}${e.code ? ` (${e.code})` : ""}`}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}
            <Field label={t("allowedHosts")} hint={locked ? t("policyLocked") : t("allowedHostsHint")}>
              <Textarea rows={4} disabled={locked} value={hostsText} onChange={(_, d) => setHostsText(d.value)} />
            </Field>
            <Caption1 className={s.muted}>
              {t("policySource")}: {policySourceLabel}
            </Caption1>
          </AccordionPanel>
        </AccordionItem>

        {/* --- Faste instruktioner pr. program (#78) --------------------------------------- */}
        <AccordionItem value="instructions" className={s.section}>
          <AccordionHeader icon={<Notepad20Regular />} size="large">
            {t("secInstructions")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            <InstructionsSection settings={settings} setSettings={setSettings} policy={policy} host={adapter?.host} />
          </AccordionPanel>
        </AccordionItem>

        {/* --- Viden fra mine mapper (F15) ------------------------------------------------ */}
        <AccordionItem value="knowledge" className={s.section}>
          <AccordionHeader icon={<BookSearch20Regular />} size="large">
            {t("secKnowledge")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            <KnowledgeSection key={saves} settings={settings} setSettings={setSettings} policy={policy} models={models} api={api.knowledge} />
          </AccordionPanel>
        </AccordionItem>

        {/* --- Excel-funktioner (#80): kun når =MIAI.ASK() er registreret ----------------- */}
        {cellFunctionsAvailable() && (
          <AccordionItem value="cells" className={s.section}>
            <AccordionHeader icon={<TableCalculator20Regular />} size="large">
              {t("secCells")}
            </AccordionHeader>
            <AccordionPanel className={s.panel}>
              <CellSection settings={settings} setSettings={setSettings} adapter={adapter} models={models} onMessage={setMsg} />
            </AccordionPanel>
          </AccordionItem>
        )}

        {/* --- Generelt ---------------------------------------------------------------- */}
        <AccordionItem value="general" className={s.section}>
          <AccordionHeader icon={<Globe20Regular />} size="large">
            {t("secGeneral")}
          </AccordionHeader>
          <AccordionPanel className={s.panel}>
            <Field label={t("uiLanguage")}>
              <Dropdown
                value={settings.uiLanguage === "da" ? "Dansk" : "English"}
                selectedOptions={[settings.uiLanguage]}
                onOptionSelect={(_, d) => setSettings({ ...settings, uiLanguage: d.optionValue as "da" | "en" })}
              >
                <Option value="da">Dansk</Option>
                <Option value="en">English</Option>
              </Dropdown>
            </Field>
            {api.mode === "bridge" && (
              <Field label={t("customTasksDir")} hint={t("customTasksDirHint")}>
                <Input value={settings.customTasksDir} onChange={(_, d) => setSettings({ ...settings, customTasksDir: d.value })} />
              </Field>
            )}
          </AccordionPanel>
        </AccordionItem>
      </Accordion>

      {/* Fast gem-bjælke nederst – altid synlig, også i lange sektioner. */}
      <div className={s.saveBar}>
        {unverified && !unverified.ok && (
          <MessageBar intent="warning">
            <MessageBarBody>
              {t(unverified.reason === "cors" ? "saveCors" : unverified.reason === "absent" ? "saveAbsent" : "saveHttp")}{" "}
              <Caption1>{unverified.message}</Caption1>
            </MessageBarBody>
            <MessageBarActions>
              <Button size="small" onClick={() => void save(true)}>
                {t("saveAnyway")}
              </Button>
              {unverified.reason === "cors" && (
                <Button size="small" onClick={() => setShowGuide(true)}>
                  {t("setupOpen")}
                </Button>
              )}
            </MessageBarActions>
          </MessageBar>
        )}
        {msg && (
          <MessageBar intent={msg.intent}>
            <MessageBarBody>{msg.text}</MessageBarBody>
          </MessageBar>
        )}
        <Button appearance="primary" icon={<Save20Regular />} onClick={() => void save()}>
          {t("save")}
        </Button>
      </div>
    </div>
  );
}

/**
 * Timeout i sekunder. Gemmes først, når feltet forlades (eller ved Enter), så man kan skrive fx 50
 * uden at "5" straks rettes til minimum; værdien holdes inden for 10–1800 s.
 */
function TimeoutInput({ valueMs, onCommit, suffix }: { valueMs: number; onCommit: (ms: number) => void; suffix: string }) {
  const [draft, setDraft] = useState(String(Math.round(valueMs / 1000)));
  useEffect(() => setDraft(String(Math.round(valueMs / 1000))), [valueMs]);
  const commit = () => {
    const sec = Math.round(Number(draft));
    const ms = (Number.isFinite(sec) && sec > 0 ? Math.min(1_800, Math.max(10, sec)) : Math.round(valueMs / 1000)) * 1000;
    setDraft(String(ms / 1000));
    if (ms !== valueMs) onCommit(ms);
  };
  return (
    <Input
      type="number"
      min={10}
      max={1800}
      step={30}
      value={draft}
      onChange={(_, d) => setDraft(d.value)}
      onBlur={commit}
      onKeyDown={(e) => e.key === "Enter" && commit()}
      contentAfter={<Caption1>{suffix}</Caption1>}
    />
  );
}
