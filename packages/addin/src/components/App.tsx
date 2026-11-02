import { useEffect, useMemo, useState } from "react";
import { Badge, Caption1, FluentProvider, Tab, TabList, Text, Tooltip, makeStyles, tokens, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { useBackend } from "../backend";
import type { HostAdapter } from "../hosts";
import { I18n, makeT, type Lang } from "../i18n";
import { ActionsTab } from "./ActionsTab";
import { ChatTab } from "./ChatTab";
import { SettingsTab } from "./SettingsTab";

const useStyles = makeStyles({
  shell: { height: "100%", display: "flex", flexDirection: "column", background: tokens.colorNeutralBackground1 },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}` },
  body: { flex: 1, overflowY: "auto", padding: tokens.spacingHorizontalM },
});

type Status = { kind: "ok" | "warn" | "down"; detail?: string };

function isDarkOffice(): boolean {
  const bg = Office.context.officeTheme?.bodyBackgroundColor;
  if (!bg || !/^#?[0-9a-f]{6}$/i.test(bg)) return window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
  const n = parseInt(bg.replace("#", ""), 16);
  const lum = 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255);
  return lum < 128;
}

export function App({ adapter }: { adapter: HostAdapter }) {
  const s = useStyles();
  const [lang, setLang] = useState<Lang>(Office.context.displayLanguage?.toLowerCase().startsWith("da") ? "da" : "en");
  const [tab, setTab] = useState<"actions" | "chat" | "settings">("actions");
  const [status, setStatus] = useState<Status>({ kind: "warn" });
  const i18n = useMemo(() => ({ lang, t: makeT(lang) }), [lang]);
  const { t } = i18n;
  const api = useBackend();

  useEffect(() => {
    api.settings().then((v) => setLang(v.settings.uiLanguage), () => {});
    const check = () =>
      api.health().then(
        (h) => setStatus(h.endpoint.reachable ? { kind: "ok", detail: h.endpoint.baseUrl } : { kind: "warn", detail: h.endpoint.error }),
        (e) => setStatus({ kind: "down", detail: (e as Error).message }),
      );
    void check();
    const id = setInterval(check, 15_000);
    return () => clearInterval(id);
  }, [api]);

  const badge = { ok: { color: "success", text: t("connected") }, warn: { color: "warning", text: t("notConnected") }, down: { color: "danger", text: t("bridgeDown") } }[status.kind] as {
    color: "success" | "warning" | "danger";
    text: string;
  };

  return (
    <I18n.Provider value={i18n}>
      {/* Bemærk: FluentProvider kopierer sin className til portaler (tooltips m.m.), så layout skal ligge på en indre div. */}
      <FluentProvider theme={isDarkOffice() ? webDarkTheme : webLightTheme} style={{ height: "100%" }}>
        <div className={s.shell}>
        <div className={s.header}>
          <Text weight="semibold">
            MIAI <Caption1>· {api.mode === "bridge" ? "bridge" : t("modeDirect")}</Caption1>
          </Text>
          <Tooltip content={status.detail ?? badge.text} relationship="description">
            <Badge appearance="tint" color={badge.color}>
              {badge.text}
            </Badge>
          </Tooltip>
        </div>
        <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as typeof tab)} size="small">
          <Tab value="actions">{t("tabActions")}</Tab>
          <Tab value="chat">{t("tabChat")}</Tab>
          <Tab value="settings">{t("tabSettings")}</Tab>
        </TabList>
        <div className={s.body}>
          {tab === "actions" && <ActionsTab adapter={adapter} />}
          {tab === "chat" && <ChatTab adapter={adapter} />}
          {tab === "settings" && <SettingsTab onLanguage={setLang} />}
        </div>
        </div>
      </FluentProvider>
    </I18n.Provider>
  );
}
