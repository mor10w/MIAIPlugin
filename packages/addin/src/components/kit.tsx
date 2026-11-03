import type { ReactNode } from "react";
import { Body1, Caption1, Skeleton, SkeletonItem, Subtitle2, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import type { FluentIcon } from "@fluentui/react-icons";
import { categoryColors, ui, type TaskCategory } from "../theme";

/**
 * Små, fælles byggeklodser til task pane (#17): ikonfelt, sektionsoverskrift, tom tilstand og
 * skelet. Alle bruger tokens og temaets konstanter, så faner ser ens ud.
 */

export const useKitStyles = makeStyles({
  tile: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
    width: ui.iconTile,
    height: ui.iconTile,
    borderRadius: tokens.borderRadiusLarge,
  },
  tileSmall: { width: ui.iconTileSmall, height: ui.iconTileSmall, borderRadius: tokens.borderRadiusMedium },
  section: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    margin: `${tokens.spacingVerticalM} 0 ${tokens.spacingVerticalXS}`,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: "0.06em",
    textTransform: "uppercase",
  },
  empty: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    textAlign: "center",
    gap: tokens.spacingVerticalS,
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalL}`,
    color: tokens.colorNeutralForeground2,
  },
  emptyIcon: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: ui.iconTile,
    height: ui.iconTile,
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    padding: tokens.spacingHorizontalS,
  },
  skeleton: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS, padding: `${tokens.spacingVerticalXS} 0` },
  /** Diskret indtoning; slås fra, når brugeren har bedt om mindre bevægelse. */
  appear: {
    "@media (prefers-reduced-motion: no-preference)": {
      animationName: { from: { opacity: 0, transform: `translateY(${ui.appearOffset})` }, to: { opacity: 1, transform: "none" } },
      animationDuration: ui.motion,
      animationTimingFunction: tokens.curveDecelerateMid,
    },
  },
});

/** Ikon i et farvet felt – farven følger opgavens kategori. */
export function IconTile({ icon: Icon, category, small }: { icon: FluentIcon; category: TaskCategory; small?: boolean }) {
  const s = useKitStyles();
  const c = categoryColors[category];
  return (
    <span className={mergeClasses(s.tile, small && s.tileSmall)} style={{ backgroundColor: c.bg, color: c.fg }} aria-hidden>
      <Icon />
    </span>
  );
}

export function SectionLabel({ children }: { children: ReactNode }) {
  return <div className={useKitStyles().section}>{children}</div>;
}

/** Tom tilstand: ikon, overskrift, én forklarende sætning og evt. næste skridt. */
export function EmptyState({ icon: Icon, title, text, action }: { icon: FluentIcon; title: string; text?: string; action?: ReactNode }) {
  const s = useKitStyles();
  return (
    <div className={mergeClasses(s.empty, s.appear)}>
      <span className={s.emptyIcon} aria-hidden>
        <Icon />
      </span>
      <Subtitle2>{title}</Subtitle2>
      {text && <Body1>{text}</Body1>}
      {action}
    </div>
  );
}

/** Skelet, mens svaret er på vej (i stedet for en spinner, når formen er kendt). */
export function TextSkeleton({ lines = 3, label }: { lines?: number; label: string }) {
  const s = useKitStyles();
  return (
    <Skeleton aria-label={label} className={s.skeleton}>
      {Array.from({ length: lines }, (_, i) => (
        <SkeletonItem key={i} size={12} style={{ width: i === lines - 1 ? "60%" : "100%" }} />
      ))}
      <Caption1>{label}</Caption1>
    </Skeleton>
  );
}
