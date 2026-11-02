import { makeStyles, tokens } from "@fluentui/react-components";
import { mdToSafeHtml } from "../markdown";

const useStyles = makeStyles({
  md: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    overflowWrap: "anywhere",
    "& p": { margin: `0 0 ${tokens.spacingVerticalS}` },
    "& h1, & h2, & h3": { margin: `${tokens.spacingVerticalM} 0 ${tokens.spacingVerticalXS}`, fontSize: tokens.fontSizeBase400 },
    "& ul, & ol": { paddingLeft: "1.2em", margin: `0 0 ${tokens.spacingVerticalS}` },
    "& code": { fontFamily: tokens.fontFamilyMonospace, background: tokens.colorNeutralBackground3, padding: "0 3px", borderRadius: "3px" },
    "& pre": { background: tokens.colorNeutralBackground3, padding: tokens.spacingHorizontalS, overflowX: "auto" },
  },
});

/** Viser modeloutput som renset HTML (se markdown.ts). */
export function Markdown({ text }: { text: string }) {
  const s = useStyles();
  return <div className={s.md} dangerouslySetInnerHTML={{ __html: mdToSafeHtml(text) }} />;
}
