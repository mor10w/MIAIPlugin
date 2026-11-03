import { useState } from "react";
import { Button, Caption1, Text, Tooltip, makeStyles, tokens } from "@fluentui/react-components";
import { ChevronDown16Regular, ChevronRight16Regular, DocumentSearch16Regular, Open16Regular } from "@fluentui/react-icons";
import type { SourceRef } from "@miai/shared";
import type { KnowledgeApi } from "../backend/types";
import { useI18n } from "../i18n";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
    marginTop: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
  },
  head: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, color: tokens.colorNeutralForeground3 },
  list: { margin: 0, padding: 0, listStyleType: "none", display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXXS },
  row: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXXS, minWidth: 0 },
  toggle: { flex: 1, minWidth: 0, justifyContent: "flex-start", textAlign: "start", fontWeight: tokens.fontWeightRegular },
  label: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  n: { fontWeight: tokens.fontWeightSemibold, color: tokens.colorBrandForeground1, marginRight: tokens.spacingHorizontalXS },
  excerpt: {
    margin: `0 0 ${tokens.spacingVerticalXS} ${tokens.spacingHorizontalXXL}`,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
  },
  path: { color: tokens.colorNeutralForeground3, overflowWrap: "anywhere" },
});

/**
 * Kilder under et chatsvar (F15): nummer som i svarets citater ([1]), filnavn, uddraget på klik
 * og "Åbn fil" (bridgen åbner filen i dens standardprogram – webviewet kan ikke åbne file://).
 */
export function SourceList({ sources, api }: { sources: SourceRef[]; api?: KnowledgeApi }) {
  const { t } = useI18n();
  const s = useStyles();
  const [open, setOpen] = useState<Record<number, string | undefined>>({});
  const [error, setError] = useState<string>();

  const toggle = async (ref: SourceRef) => {
    if (ref.n in open) {
      setOpen(({ [ref.n]: _, ...rest }) => rest);
      return;
    }
    setOpen((o) => ({ ...o, [ref.n]: ref.snippet }));
    // Hele uddraget hentes, hvis bridgen stadig har det (indekset kan være opdateret siden).
    const full = await api?.chunk(ref.chunkId).catch(() => undefined);
    if (full) setOpen((o) => (ref.n in o ? { ...o, [ref.n]: full.text } : o));
  };

  const openFile = async (ref: SourceRef) => {
    setError(undefined);
    await api?.open(ref.chunkId).catch((err: Error) => setError(err.message));
  };

  if (!sources.length) return null;
  return (
    <div className={s.root}>
      <Caption1 className={s.head}>
        <DocumentSearch16Regular aria-hidden />
        {t("sourcesTitle")}
      </Caption1>
      <ul className={s.list}>
        {sources.map((ref) => {
          const expanded = ref.n in open;
          return (
            <li key={ref.n}>
              <div className={s.row}>
                <Button
                  size="small"
                  appearance="subtle"
                  className={s.toggle}
                  icon={expanded ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
                  aria-expanded={expanded}
                  onClick={() => void toggle(ref)}
                >
                  <span className={s.label}>
                    <span className={s.n}>[{ref.n}]</span>
                    {ref.title}
                  </span>
                </Button>
                {api && (
                  <Tooltip content={t("sourceOpen")} relationship="label">
                    <Button size="small" appearance="subtle" icon={<Open16Regular />} onClick={() => void openFile(ref)} />
                  </Tooltip>
                )}
              </div>
              {expanded && (
                <div className={s.excerpt}>
                  <Text size={200}>{open[ref.n]}</Text>
                  <Caption1 block className={s.path}>
                    {ref.path}
                  </Caption1>
                </div>
              )}
            </li>
          );
        })}
      </ul>
      {error && <Caption1 className={s.path}>{error}</Caption1>}
    </div>
  );
}
