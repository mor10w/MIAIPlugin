import { useCallback, useEffect, useState } from "react";
import {
  Button,
  Caption1,
  Combobox,
  Field,
  Input,
  Option,
  ProgressBar,
  SpinButton,
  Switch,
  Text,
  Tooltip,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ArrowSync16Regular, Delete20Regular, FolderAdd20Regular } from "@fluentui/react-icons";
import type { KnowledgeStatus, Policy, Settings } from "@miai/shared";
import type { KnowledgeApi } from "../backend/types";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  muted: { color: tokens.colorNeutralForeground3 },
  folders: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS, margin: 0, padding: 0, listStyleType: "none" },
  folder: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, minWidth: 0 },
  folderText: { flex: 1, minWidth: 0, overflowWrap: "anywhere", fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
  addRow: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "flex-end" },
  grow: { flex: 1, minWidth: 0 },
  status: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  statusHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  error: { color: tokens.colorPaletteRedForeground1, overflowWrap: "anywhere" },
});

interface Props {
  settings: Settings;
  setSettings: (s: Settings) => void;
  policy: Policy;
  /** Modeller på den aktive endpoint – embedding-modeller foreslås først. */
  models: string[];
  /** Udeladt i direkte tilstand: browseren kan ikke læse brugerens mapper. */
  api?: KnowledgeApi;
}

/** Indstillinger for lokal RAG (F15, ADR-0013): mapper, embedding-model og indeksets status. */
export function KnowledgeSection({ settings, setSettings, policy, models, api }: Props) {
  const { t, lang } = useI18n();
  const s = useStyles();
  const k = settings.knowledge;
  const set = (patch: Partial<Settings["knowledge"]>) => setSettings({ ...settings, knowledge: { ...k, ...patch } });
  const [folder, setFolder] = useState("");
  const [status, setStatus] = useState<KnowledgeStatus>();
  const [error, setError] = useState<string>();

  const refresh = useCallback(() => api?.status().then(setStatus, (e: Error) => setError(e.message)), [api]);
  useEffect(() => {
    void refresh();
  }, [refresh]);
  // Under indeksering opdateres status løbende.
  useEffect(() => {
    if (status?.state !== "indexing") return;
    const id = setInterval(() => void refresh(), 1_000);
    return () => clearInterval(id);
  }, [status?.state, refresh]);

  if (!api) return <Caption1 className={s.muted}>{t("knowledgeBridgeOnly")}</Caption1>;

  const addFolder = () => {
    const f = folder.trim();
    if (!f || k.folders.includes(f)) return;
    set({ folders: [...k.folders, f] });
    setFolder("");
  };
  const embedFirst = [...models].sort((a, b) => Number(/embed|bge|e5|gte|minilm/i.test(b)) - Number(/embed|bge|e5|gte|minilm/i.test(a)));
  const reindex = async () => {
    setError(undefined);
    setStatus(await api.reindex().catch((e: Error) => (setError(e.message), status)));
  };

  return (
    <div className={s.root}>
      <Caption1 className={s.muted}>{t("knowledgeIntro")}</Caption1>
      <Switch
        label={t("knowledgeEnable")}
        disabled={!policy.knowledge}
        checked={k.enabled && policy.knowledge}
        onChange={(_, d) => set({ enabled: d.checked })}
      />

      {k.enabled && policy.knowledge && (
        <>
          <Field label={t("knowledgeFolders")}>
            <ul className={s.folders}>
              {k.folders.map((f) => (
                <li key={f} className={s.folder}>
                  <Text className={s.folderText}>{f}</Text>
                  <Tooltip content={t("knowledgeRemoveFolder")} relationship="label">
                    <Button size="small" appearance="subtle" icon={<Delete20Regular />} onClick={() => set({ folders: k.folders.filter((x) => x !== f) })} />
                  </Tooltip>
                </li>
              ))}
            </ul>
          </Field>
          <div className={s.addRow}>
            <Field className={s.grow} label={t("knowledgeAddFolder")}>
              <Input
                value={folder}
                placeholder={t("knowledgeFolderPlaceholder")}
                onChange={(_, d) => setFolder(d.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") addFolder();
                }}
              />
            </Field>
            <Tooltip content={t("knowledgeAddFolder")} relationship="label">
              <Button icon={<FolderAdd20Regular />} disabled={!folder.trim()} onClick={addFolder} />
            </Tooltip>
          </div>

          <Field label={t("knowledgeModel")} hint={t("knowledgeModelHint")}>
            <Combobox
              freeform
              value={k.embeddingModel}
              selectedOptions={[k.embeddingModel]}
              onOptionSelect={(_, d) => set({ embeddingModel: d.optionValue ?? "" })}
              onChange={(e) => set({ embeddingModel: (e.target as HTMLInputElement).value })}
            >
              {embedFirst.map((m) => (
                <Option key={m} value={m}>
                  {m}
                </Option>
              ))}
            </Combobox>
          </Field>

          <Field label={t("knowledgeTopK")}>
            <SpinButton min={1} max={12} value={k.topK} onChange={(_, d) => set({ topK: Number(d.value ?? d.displayValue ?? k.topK) || k.topK })} />
          </Field>

          {status && (
            <div className={s.status} aria-live="polite">
              <div className={s.statusHead}>
                <Text weight="semibold">{t("knowledgeStatus", { documents: status.documents, chunks: status.chunks })}</Text>
                <Button size="small" icon={<ArrowSync16Regular />} disabled={status.state === "indexing"} onClick={() => void reindex()}>
                  {t("knowledgeReindex")}
                </Button>
              </div>
              {status.state === "indexing" && (
                <>
                  <ProgressBar value={status.total ? status.done / status.total : undefined} />
                  <Caption1>{t("knowledgeIndexing", { done: status.done, total: status.total })}</Caption1>
                </>
              )}
              {status.lastIndexed && status.state !== "indexing" && (
                <Caption1 className={s.muted}>
                  {t("knowledgeLast", { when: new Date(status.lastIndexed).toLocaleString(lang === "da" ? "da-DK" : "en-GB") })}
                </Caption1>
              )}
              {status.skippedFolders.map((f) => (
                <Caption1 key={f.folder} className={s.error}>
                  {t(f.reason === "policy" ? "knowledgeSkippedPolicy" : "knowledgeSkippedMissing", { folder: f.folder })}
                </Caption1>
              ))}
              {status.failed.length > 0 && <Caption1 className={s.muted}>{t("knowledgeFailed", { n: status.failed.length })}</Caption1>}
              {status.error && <Caption1 className={s.error}>{status.error}</Caption1>}
              {status.state === "disabled" && <Caption1 className={s.muted}>{t("knowledgeSaveFirst")}</Caption1>}
            </div>
          )}
          {error && <Caption1 className={s.error}>{error}</Caption1>}
        </>
      )}
    </div>
  );
}
