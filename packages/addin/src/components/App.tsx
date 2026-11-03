import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Badge,
  Button,
  Caption1,
  FluentProvider,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Tab,
  TabList,
  Text,
  Tooltip,
  makeStyles,
  tokens,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { useBackend } from "../backend";
import type { HostAdapter } from "../hosts";
import { I18n, makeT, type Lang } from "../i18n";
import { ActionsTab } from "./ActionsTab";
import { ChatTab } from "./ChatTab";
import { SettingsTab } from "./SettingsTab";
import { TasksTab } from "./TasksTab";

const useStyles = makeStyles({
  shell: { height: "100%", display: "flex", flexDirection: "column", background: tokens.colorNeutralBackground1 },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}` },
  body: { flex: 1, overflowY: "auto", padding: tokens.spacingHorizontalM },
  banner: { marginBottom: tokens.spacingVerticalM },
});

/** token = bridgen afviser, fordi adgangstokenet mangler (gammelt manifest, #11). */
type Status = { kind: "ok" | "warn" | "down"; detail?: string; token?: boolean };

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
  const [tab, setTab] = useState<"actions" | "tasks" | "chat" | "settings">("actions");
  const [status, setStatus] = useState<Status>({ kind: "warn" });
  /** Først når første statustjek er færdigt, vises "Kom i gang" (ellers blinker banneret ved start). */
  const [checked, setChecked] = useState(false);
  const i18n = useMemo(() => ({ lang, t: makeT(lang) }), [lang]);
  const { t } = i18n;
  const api = useBackend();

  const check = useCallback(
    () =>
      api
        .health()
        .then(
          (h) => setStatus(h.endpoint.reachable ? { kind: "ok", detail: h.endpoint.baseUrl } : { kind: "warn", detail: h.endpoint.error }),
          (e) => setStatus({ kind: "down", detail: (e as Error).message, token: (e as { code?: string }).code === "TOKEN" }),
        )
        .finally(() => setChecked(true)),
    [api],
  );

  useEffect(() => {
    api.settings().then(
      (v) => setLang(v.settings.uiLanguage),
      () => {},
    );
    void check();
    const id = setInterval(() => void check(), 15_000);
    return () => clearInterval(id);
  }, [api, check]);

  const badge = {
    ok: { color: "success", text: t("connected") },
    warn: { color: "warning", text: t("notConnected") },
    down: { color: "danger", text: t("bridgeDown") },
  }[status.kind] as {
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
            <Tab value="tasks">{t("tabTasks")}</Tab>
            <Tab value="chat">{t("tabChat")}</Tab>
            <Tab value="settings">{t("tabSettings")}</Tab>
          </TabList>
          <div className={s.body}>
            {/* Førstegangsguide (#4): uden forbindelse peges brugeren til Indstillinger, hvor guiden åbner. */}
            {status.token && (
              <MessageBar intent="error" className={s.banner}>
                <MessageBarBody>
                  <MessageBarTitle>{t("tokenTitle")}</MessageBarTitle>
                  {t("tokenHelp")}
                </MessageBarBody>
              </MessageBar>
            )}
            {checked && status.kind !== "ok" && !status.token && tab !== "settings" && (
              <MessageBar intent="warning" className={s.banner}>
                <MessageBarBody>{t("noConnectionBanner")}</MessageBarBody>
                <MessageBarActions>
                  <Button size="small" appearance="primary" onClick={() => setTab("settings")}>
                    {t("getStarted")}
                  </Button>
                </MessageBarActions>
              </MessageBar>
            )}
            {tab === "actions" && <ActionsTab adapter={adapter} />}
            {tab === "tasks" && <TasksTab />}
            {tab === "chat" && <ChatTab adapter={adapter} />}
            {tab === "settings" && <SettingsTab onLanguage={setLang} onSaved={() => void check()} />}
          </div>
        </div>
      </FluentProvider>
    </I18n.Provider>
  );
}
