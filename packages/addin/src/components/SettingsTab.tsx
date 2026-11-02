import { useEffect, useState } from "react";
import {
  Button,
  Caption1,
  Card,
  Dropdown,
  Field,
  Input,
  MessageBar,
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
import type { Endpoint, Policy, Settings, StructuredOutputMode } from "@miai/shared";
import { corsHint, useBackend } from "../backend";
import { useI18n } from "../i18n";

/** Kendte OpenAI-kompatible inference-værktøjer og deres standardporte. "Andet" dækker alt andet. */
const PRESETS: { id: string; name: string; baseUrl: string; structuredOutput: StructuredOutputMode }[] = [
  { id: "ollama", name: "Ollama", baseUrl: "http://127.0.0.1:11434/v1", structuredOutput: "json_schema" },
  { id: "lmstudio", name: "LM Studio", baseUrl: "http://127.0.0.1:1234/v1", structuredOutput: "json_schema" },
  { id: "llamacpp", name: "llama.cpp server", baseUrl: "http://127.0.0.1:8080/v1", structuredOutput: "json_schema" },
  { id: "vllm", name: "vLLM", baseUrl: "http://127.0.0.1:8000/v1", structuredOutput: "json_schema" },
  { id: "localai", name: "LocalAI", baseUrl: "http://127.0.0.1:8080/v1", structuredOutput: "json_object" },
  { id: "jan", name: "Jan", baseUrl: "http://127.0.0.1:1337/v1", structuredOutput: "json_object" },
  { id: "kobold", name: "KoboldCpp", baseUrl: "http://127.0.0.1:5001/v1", structuredOutput: "none" },
  { id: "custom", name: "Andet / Other", baseUrl: "http://127.0.0.1:8000/v1", structuredOutput: "none" },
];

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  card: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  row: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center", flexWrap: "wrap" },
});

function uniqueId(base: string, taken: string[]): string {
  let id = base;
  for (let i = 2; taken.includes(id); i++) id = `${base}-${i}`;
  return id;
}

export function SettingsTab({ onLanguage }: { onLanguage: (l: "da" | "en") => void }) {
  const { t } = useI18n();
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
    api.models().then((m) => setModels(m.map((x) => x.id)), () => setModels([]));
  }, [api]);

  if (!settings || !policy) return msg ? <MessageBar intent="error"><MessageBarBody>{msg.text}</MessageBarBody></MessageBar> : <Spinner />;

  const locked = policy.locked || policySource === "machine";
  const setEp = (id: string, patch: Partial<Endpoint>) => setSettings({ ...settings, endpoints: settings.endpoints.map((e) => (e.id === id ? { ...e, ...patch } : e)) });

  const addEndpoint = () => {
    const p = PRESETS.find((x) => x.id === preset)!;
    const id = uniqueId(p.id, settings.endpoints.map((e) => e.id));
    const ep: Endpoint = { id, name: p.name, baseUrl: p.baseUrl, apiKey: "", structuredOutput: p.structuredOutput, stripThinking: true, extraBody: {}, timeoutMs: 120_000 };
    setSettings({ ...settings, endpoints: [...settings.endpoints, ep] });
  };

  const save = async () => {
    setMsg(undefined);
    const nextPolicy = locked ? undefined : { ...policy, allowedHosts: hostsText.split(/\s+/).map((h) => h.trim()).filter(Boolean) };
    try {
      await api.saveSettings(settings, nextPolicy);
      if (nextPolicy) setPolicy(nextPolicy);
      onLanguage(settings.uiLanguage);
      setMsg({ intent: "success", text: t("saved") });
    } catch (e) {
      setMsg({ intent: "error", text: (e as Error).message });
    }
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
        <MessageBarBody>{api.mode === "bridge" ? t("modeBridgeInfo") : `${t("modeDirectInfo")} ${corsHint()}`}</MessageBarBody>
      </MessageBar>
      <Subtitle2>{t("endpoints")}</Subtitle2>
      {settings.endpoints.map((ep) => (
        <Card key={ep.id} className={s.card}>
          <div className={s.row}>
            <Radio label={t("active")} checked={settings.activeEndpoint === ep.id} onChange={() => setSettings({ ...settings, activeEndpoint: ep.id, defaultModel: "" })} />
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
            <Dropdown value={ep.structuredOutput} selectedOptions={[ep.structuredOutput]} onOptionSelect={(_, d) => setEp(ep.id, { structuredOutput: d.optionValue as StructuredOutputMode })}>
              <Option value="json_schema">json_schema</Option>
              <Option value="json_object">json_object</Option>
              <Option value="none">none</Option>
            </Dropdown>
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
      <div className={s.row}>
        <Field label={t("temperature")}>
          <Input type="number" min={0} max={2} step={0.1} value={String(settings.temperature)} onChange={(_, d) => setSettings({ ...settings, temperature: Number(d.value) })} />
        </Field>
        <Field label={t("maxTokens")}>
          <Input type="number" min={16} step={256} value={String(settings.maxTokens)} onChange={(_, d) => setSettings({ ...settings, maxTokens: Number(d.value) })} />
        </Field>
      </div>
      <Field label={t("uiLanguage")}>
        <Dropdown value={settings.uiLanguage === "da" ? "Dansk" : "English"} selectedOptions={[settings.uiLanguage]} onOptionSelect={(_, d) => setSettings({ ...settings, uiLanguage: d.optionValue as "da" | "en" })}>
          <Option value="da">Dansk</Option>
          <Option value="en">English</Option>
        </Dropdown>
      </Field>
      <Switch label={t("auditLog")} checked={settings.auditLog} onChange={(_, d) => setSettings({ ...settings, auditLog: d.checked })} />
      <Field label={t("allowedHosts")} hint={locked ? t("policyLocked") : undefined}>
        <Textarea rows={4} disabled={locked} value={hostsText} onChange={(_, d) => setHostsText(d.value)} />
      </Field>
      <Caption1>
        {t("policySource")}: {policySource}
      </Caption1>

      <Button appearance="primary" onClick={save}>
        {t("save")}
      </Button>
      {msg && (
        <MessageBar intent={msg.intent}>
          <MessageBarBody>{msg.text}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}
