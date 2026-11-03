import { useEffect, useState } from "react";
import {
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
  Subtitle2,
  Switch,
  Textarea,
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
  type Settings,
  type StructuredOutputMode,
} from "@miai/shared";
import { useBackend } from "../backend";
import { useI18n } from "../i18n";
import { SetupGuide } from "./SetupGuide";
import { TaskModels } from "./TaskModels";

/** Kendte værktøjer (fælles liste i shared, #4) + "Andet" til alt andet. */
const PRESETS: KnownTool[] = [...KNOWN_TOOLS, { id: "custom", name: "Andet / Other", baseUrl: "http://127.0.0.1:8000/v1", structuredOutput: "none" }];

/** presence_penalty-forudindstillinger. "" = sendes ikke (serverens standard). */
const PENALTY_PRESETS = [
  { value: "", label: "penaltyOff" },
  { value: "0.5", label: "penaltyLight" },
  { value: "1", label: "penaltyMedium" },
  { value: "1.5", label: "penaltyStrong" },
] as const;

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  row: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center", flexWrap: "wrap" },
  logCard: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  logWrap: { overflowX: "auto", maxHeight: "320px", overflowY: "auto" },
  logTable: {
    borderCollapse: "collapse",
    fontSize: tokens.fontSizeBase200,
    width: "100%",
    "& td, & th": { border: `1px solid ${tokens.colorNeutralStroke2}`, padding: `2px ${tokens.spacingHorizontalXS}`, whiteSpace: "nowrap", textAlign: "left" },
  },
});

function uniqueId(base: string, taken: string[]): string {
  let id = base;
  for (let i = 2; taken.includes(id); i++) id = `${base}-${i}`;
  return id;
}

export function SettingsTab({ onLanguage, onSaved }: { onLanguage: (l: "da" | "en") => void; onSaved?: () => void }) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const api = useBackend();
  const [settings, setSettings] = useState<Settings>();
  const [policy, setPolicy] = useState<Policy>();
  const [policySource, setPolicySource] = useState<string>("default");
  const [hostsText, setHostsText] = useState("");
  const [models, setModels] = useState<string[]>([]);
  const [msg, setMsg] = useState<{ intent: "success" | "error" | "info"; text: string }>();
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

  /** Viser også en værdi, der er sat manuelt i config.json og ikke er en af forudindstillingerne. */
  const penaltyLabel = (v: number | undefined) => {
    const preset = PENALTY_PRESETS.find((p) => (p.value ? Number(p.value) : undefined) === v);
    return preset ? t(preset.label) : String(v);
  };
  const locked = policy.locked || policySource === "machine";
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
      timeoutMs: 120_000,
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
      if (nextPolicy) setPolicy(nextPolicy);
      onLanguage(next.uiLanguage);
      onSaved?.();
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
      timeoutMs: 120_000,
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
      await api.saveSettings(settings, locked ? undefined : { ...policy, allowedHosts: hostsText.split(/\s+/).filter(Boolean) });
      const m = await api.models();
      setModels(m.map((x) => x.id));
      setMsg({ intent: "success", text: t("modelsFound", { n: m.length }) });
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    } finally {
      setTesting(false);
    }
  };

  return (
    <div className={s.root}>
      <MessageBar intent="info">
        {/* CORS-vejledningen står præcist pr. værktøj i guiden nedenfor (#4). */}
        <MessageBarBody>{api.mode === "bridge" ? t("modeBridgeInfo") : t("modeDirectInfo")}</MessageBarBody>
      </MessageBar>
      {showGuide ? (
        <SetupGuide onUse={useTool} />
      ) : (
        <div>
          <Button size="small" onClick={() => setShowGuide(true)}>
            {t("setupOpen")}
          </Button>
        </div>
      )}
      <Subtitle2>{t("endpoints")}</Subtitle2>
      {settings.endpoints.map((ep) => (
        <Card key={ep.id} className={s.card}>
          <div className={s.row}>
            <Radio
              label={t("active")}
              checked={settings.activeEndpoint === ep.id}
              onChange={() => setSettings({ ...settings, activeEndpoint: ep.id, defaultModel: "" })}
            />
            {settings.endpoints.length > 1 && (
              <Button
                size="small"
                appearance="subtle"
                onClick={() => {
                  const rest = settings.endpoints.filter((e) => e.id !== ep.id);
                  setSettings({ ...settings, endpoints: rest, activeEndpoint: settings.activeEndpoint === ep.id ? rest[0].id : settings.activeEndpoint });
                }}
              >
                {t("remove")}
              </Button>
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
          <Field label={t("structuredOutput")}>
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
          <Switch label={t("stripThinking")} checked={ep.stripThinking} onChange={(_, d) => setEp(ep.id, { stripThinking: d.checked })} />
        </Card>
      ))}
      <div className={s.row}>
        <Dropdown value={PRESETS.find((p) => p.id === preset)?.name} selectedOptions={[preset]} onOptionSelect={(_, d) => setPreset(d.optionValue ?? "custom")}>
          {PRESETS.map((p) => (
            <Option key={p.id} value={p.id}>
              {p.name}
            </Option>
          ))}
        </Dropdown>
        <Button onClick={addEndpoint}>{t("addEndpoint")}</Button>
      </div>

      <div className={s.row}>
        <Button onClick={test} disabled={testing}>
          {t("testConnection")}
        </Button>
        {testing && <Spinner size="tiny" />}
      </div>

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
      <TaskModels value={settings.taskModels} models={models} onChange={(taskModels) => setSettings({ ...settings, taskModels })} />
      <div className={s.row}>
        <Field label={t("temperature")}>
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
      <Switch label={t("auditLog")} checked={settings.auditLog} onChange={(_, d) => setSettings({ ...settings, auditLog: d.checked })} />
      {api.mode === "bridge" && (
        <Field label={t("customTasksDir")} hint={t("customTasksDirHint")}>
          <Input value={settings.customTasksDir} onChange={(_, d) => setSettings({ ...settings, customTasksDir: d.value })} />
        </Field>
      )}
      <div className={s.row}>
        <Button onClick={() => void toggleAudit()}>{showAudit ? t("hideAudit") : t("viewAudit")}</Button>
      </div>
      {showAudit && (
        <Card className={s.logCard}>
          {!auditEntries ? (
            <Spinner />
          ) : auditEntries.length === 0 ? (
            t("auditEmpty")
          ) : (
            <div className={s.logWrap}>
              <table className={s.logTable}>
                <thead>
                  <tr>
                    <th>{t("auditTime")}</th>
                    <th>{t("auditTask")}</th>
                    <th>host</th>
                    <th>{t("model")}</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {auditEntries.map((e, i) => (
                    <tr key={i}>
                      <td>{new Date(e.ts).toLocaleString(lang === "da" ? "da-DK" : "en-GB")}</td>
                      <td>{e.kind === "task" ? e.task : "models"}</td>
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
        </Card>
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
      {/* Søgetjeneste (#12): DuckDuckGo eller selvhostet SearXNG; kan låses af organisationens politik. */}
      {api.mode === "bridge" && policy.webSearch && settings.webSearch && (
        <div className={s.row}>
          <Field label={t("searchProvider")} hint={policy.search ? t("searchLocked") : undefined}>
            <Dropdown
              disabled={!!policy.search}
              value={(policy.search ?? settings.search).provider === "searxng" ? "SearXNG" : "DuckDuckGo"}
              selectedOptions={[(policy.search ?? settings.search).provider]}
              onOptionSelect={(_, d) => setSettings({ ...settings, search: { ...settings.search, provider: d.optionValue as "duckduckgo" | "searxng" } })}
            >
              <Option value="duckduckgo">DuckDuckGo</Option>
              <Option value="searxng">SearXNG</Option>
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
      <Field label={t("allowedHosts")} hint={locked ? t("policyLocked") : undefined}>
        <Textarea rows={4} disabled={locked} value={hostsText} onChange={(_, d) => setHostsText(d.value)} />
      </Field>
      <Caption1>
        {t("policySource")}: {policySource}
      </Caption1>

      <Button appearance="primary" onClick={() => void save()}>
        {t("save")}
      </Button>
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
    </div>
  );
}
