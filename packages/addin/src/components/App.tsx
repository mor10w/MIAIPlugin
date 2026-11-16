import { useCallback, useEffect, useMemo, useState } from "react";
import { policyReloadPending } from "../csp";
import {
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
  mergeClasses,
  tokens,
} from "@fluentui/react-components";
import { Chat20Regular, Flash20Regular, PuzzlePiece20Regular, Settings20Regular, type FluentIcon } from "@fluentui/react-icons";
import { useBackend } from "../backend";
import type { HostAdapter } from "../hosts";
import { I18n, makeT, type Lang } from "../i18n";
import { statusColors, ui, useOfficeTheme } from "../theme";
import { ActionsTab } from "./ActionsTab";
import { ChatTab, type AgentRequest } from "./ChatTab";
import { useKitStyles } from "./kit";
import { SettingsTab } from "./SettingsTab";
import { TasksTab } from "./TasksTab";

const useStyles = makeStyles({
  shell: { height: "100%", display: "flex", flexDirection: "column", backgroundColor: tokens.colorNeutralBackground2, color: tokens.colorNeutralForeground1 },
  top: { backgroundColor: tokens.colorNeutralBackground1, borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}` },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM} 0`,
  },
  brand: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, minWidth: 0 },
  brandText: { display: "flex", flexDirection: "column", lineHeight: tokens.lineHeightBase200 },
  mode: { color: tokens.colorNeutralForeground3 },
  status: { display: "inline-flex", alignItems: "center", gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground2, minWidth: 0 },
  dot: { width: ui.dot, height: ui.dot, borderRadius: tokens.borderRadiusCircular, flexShrink: 0 },
  tabs: { padding: `0 ${tokens.spacingHorizontalXS}` },
  /**
   * Smalt panel: kun den valgte fane viser tekst, de andre kun ikon (navnet står i aria-label og
   * tooltip). Allersmallest vises kun ikoner.
   */
  tabTextOther: { [ui.mqNarrow]: { display: "none" } },
  tabTextSelected: { [ui.mqTiny]: { display: "none" } },
  body: { flex: 1, overflowY: "auto", overflowX: "hidden", padding: tokens.spacingHorizontalM, minWidth: 0 },
  pane: { height: "100%" },
  banner: { marginBottom: tokens.spacingVerticalM },
});

/** token = bridgen afviser, fordi adgangstokenet mangler (gammelt manifest, #11). */
type Status = { kind: "ok" | "warn" | "down"; detail?: string; token?: boolean };
type TabId = "actions" | "tasks" | "chat" | "settings";

/** MIAI's mærke: en firtakket gnist i et afrundet felt i brandfarven (kun tokens – følger temaet). */
function LogoMark() {
  return (
    <svg width={ui.iconTile} height={ui.iconTile} viewBox="0 0 32 32" aria-hidden focusable="false">
      <rect width="32" height="32" rx="9" style={{ fill: tokens.colorBrandBackground }} />
      <path
        d="M16 6c.7 4.6 2.4 7.3 7.7 8.6.6.2.6 1 0 1.2-5.3 1.3-7 4-7.7 8.6-.1.7-1 .7-1.1 0-.7-4.6-2.4-7.3-7.7-8.6-.6-.2-.6-1 0-1.2 5.3-1.3 7-4 7.7-8.6.1-.7 1-.7 1.1 0Z"
        style={{ fill: tokens.colorNeutralForegroundOnBrand }}
      />
      <circle cx="24" cy="8.5" r="1.8" style={{ fill: tokens.colorNeutralForegroundOnBrand, opacity: 0.85 }} />
    </svg>
  );
}

const TABS: { id: TabId; icon: FluentIcon; label: "tabActions" | "tabTasks" | "tabChat" | "tabSettings" }[] = [
  { id: "actions", icon: Flash20Regular, label: "tabActions" },
  { id: "tasks", icon: PuzzlePiece20Regular, label: "tabTasks" },
  { id: "chat", icon: Chat20Regular, label: "tabChat" },
  { id: "settings", icon: Settings20Regular, label: "tabSettings" },
];

export function App({ adapter }: { adapter: HostAdapter }) {
  const s = useStyles();
  const kit = useKitStyles();
  const theme = useOfficeTheme();
  const [lang, setLang] = useState<Lang>(Office.context.displayLanguage?.toLowerCase().startsWith("da") ? "da" : "en");
  // Efter en genindlæsning, fordi politikkens hosts blev ændret (#28), åbnes Indstillinger igen.
  const [tab, setTab] = useState<TabId>(() => (policyReloadPending() ? "settings" : "actions"));
  const [chatOpened, setChatOpened] = useState(false);
  if (tab === "chat" && !chatOpened) setChatOpened(true);
  /** Prompt fra chatten → udfyldes i agent-opgaven, når brugeren vælger "Gør det med agenten". */
  const [agentRequest, setAgentRequest] = useState<AgentRequest | null>(null);
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

  const statusText = { ok: t("connected"), warn: t("notConnected"), down: t("bridgeDown") }[status.kind];

  return (
    <I18n.Provider value={i18n}>
      {/* Bemærk: FluentProvider kopierer sin className til portaler (tooltips m.m.), så layout skal ligge på en indre div. */}
      <FluentProvider theme={theme} style={{ height: "100%" }}>
        <div className={s.shell}>
          <div className={s.top}>
            <header className={s.header}>
              <div className={s.brand}>
                <LogoMark />
                <div className={s.brandText}>
                  <Text weight="semibold">MIAI</Text>
                  <Caption1 className={s.mode}>{api.mode === "bridge" ? t("modeBridge") : t("modeDirect")}</Caption1>
                </div>
              </div>
              <Tooltip content={status.detail ?? statusText} relationship="description">
                <Button
                  appearance="transparent"
                  size="small"
                  className={s.status}
                  onClick={() => status.kind !== "ok" && setTab("settings")}
                  aria-label={`${t("statusLabel")}: ${statusText}`}
                >
                  <span className={s.dot} style={{ backgroundColor: statusColors[status.kind] }} aria-hidden />
                  <Caption1>{statusText}</Caption1>
                </Button>
              </Tooltip>
            </header>
            <TabList className={s.tabs} selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as TabId)} size="small">
              {TABS.map(({ id, icon: Icon, label }) => {
                const item = (
                  <Tab key={id} value={id} icon={<Icon />} aria-label={t(label)}>
                    <span className={id === tab ? s.tabTextSelected : s.tabTextOther}>{t(label)}</span>
                  </Tab>
                );
                // Den valgte fane viser selv sin tekst; de andre kan være ikon-only i et smalt panel.
                return id === tab ? (
                  item
                ) : (
                  <Tooltip key={id} content={t(label)} relationship="description" positioning="below">
                    {item}
                  </Tooltip>
                );
              })}
            </TabList>
          </div>
          <main className={s.body}>
            {status.token && (
              <MessageBar intent="error" className={s.banner}>
                <MessageBarBody>
                  <MessageBarTitle>{t("tokenTitle")}</MessageBarTitle>
                  {t("tokenHelp")}
                </MessageBarBody>
              </MessageBar>
            )}
            {/* Førstegangsguide (#4): uden forbindelse peges brugeren til Indstillinger, hvor guiden åbner. */}
            {checked && status.kind !== "ok" && !status.token && tab !== "settings" && (
              <MessageBar intent="warning" className={mergeClasses(s.banner, kit.appear)}>
                <MessageBarBody>{t("noConnectionBanner")}</MessageBarBody>
                <MessageBarActions>
                  <Button size="small" appearance="primary" onClick={() => setTab("settings")}>
                    {t("getStarted")}
                  </Button>
                </MessageBarActions>
              </MessageBar>
            )}
            {tab === "actions" && <ActionsTab adapter={adapter} agentRequest={agentRequest} onAgentRequestConsumed={() => setAgentRequest(null)} />}
            {tab === "tasks" && <TasksTab />}
            {/* Chatten bygges første gang, den åbnes, og bliver derefter liggende skjult, så en samtale
                i gang (et svar, der skrives, en plan, kladden) ikke går tabt ved et kort faneskift. */}
            {chatOpened && (
              <div className={s.pane} hidden={tab !== "chat"}>
                <ChatTab
                  adapter={adapter}
                  active={tab === "chat"}
                  onUseAgent={(p) => {
                    setAgentRequest(p);
                    setTab("actions");
                  }}
                />
              </div>
            )}
            {tab === "settings" && <SettingsTab onLanguage={setLang} onSaved={() => void check()} />}
          </main>
        </div>
      </FluentProvider>
    </I18n.Provider>
  );
}
