import { useEffect, useRef } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { mdToSafeHtml } from "../markdown";
import { mermaidToSvg } from "../mermaid";
import { sanitizeSvg } from "../svg";

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
    "& .figure": { margin: `0 0 ${tokens.spacingVerticalS}`, overflowX: "auto" },
    "& .figure img": { maxWidth: "100%", height: "auto" },
  },
});

/** SVG → data-url (base64). Sikrer, at non-ASCII tegn overlever btoa. */
function svgToDataUrl(svg: string): string {
  const bytes = new TextEncoder().encode(svg);
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return `data:image/svg+xml;base64,${btoa(bin)}`;
}

/**
 * Viser modeloutput som renset HTML (se markdown.ts). ```mermaid-blokke renderes
 * visuelt via mermaid.js, og ```svg-blokke vises som billede (data-url – SVG'en
 * køres ikke, og renses først). Ved fejl forbliver kodeblokken synlig.
 */
export function Markdown({ text }: { text: string }) {
  const s = useStyles();
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const root = ref.current;
    if (!root) return;
    let cancelled = false;

    const replaceBlock = (block: Element, node: Node) => {
      if (!block.isConnected) return;
      // Erstat kun <pre>-blokken – ikke hele beskeden (parentElement.parentElement var Markdown-roden).
      const pre = block.closest("pre");
      if (pre && root.contains(pre)) pre.replaceWith(node);
    };

    const renderMermaid = async () => {
      for (const block of Array.from(root.querySelectorAll("code.language-mermaid"))) {
        try {
          const svg = await mermaidToSvg(block.textContent ?? "");
          if (cancelled) return;
          const wrap = document.createElement("div");
          wrap.className = "figure";
          // SVG'en kommer fra mermaid.js (securityLevel: strict), ikke fra modeloutput.
          wrap.innerHTML = svg;
          replaceBlock(block, wrap);
        } catch {
          // Ugyldigt diagram: vis kodeblokken som den er.
        }
      }
    };

    const renderSvg = () => {
      for (const block of Array.from(root.querySelectorAll("code.language-svg"))) {
        try {
          const svg = sanitizeSvg(block.textContent ?? ""); // modeloutput: rens altid først
          if (!svg.trim()) throw new Error("tom blok");
          const img = document.createElement("img");
          // Data-url: browseren behandler SVG'en som billede – ingen scripts, ingen eksterne referencer.
          img.src = svgToDataUrl(svg);
          img.alt = "Figur";
          const wrap = document.createElement("div");
          wrap.className = "figure";
          wrap.appendChild(img);
          replaceBlock(block, wrap);
        } catch {
          // Ugyldig SVG: vis kodeblokken som den er.
        }
      }
    };

    renderSvg();
    // Mermaid: vent på en pause i strømmen, så et halvfærdigt diagram ikke rendres ved hvert token.
    // mermaidToSvg cacher pr. kode, så færdige diagrammer vises straks ved genrendering.
    const timer = root.querySelector("code.language-mermaid") ? setTimeout(() => void renderMermaid(), 300) : undefined;

    return () => {
      cancelled = true;
      if (timer) clearTimeout(timer);
    };
  }, [text]);

  return <div ref={ref} className={s.md} dangerouslySetInnerHTML={{ __html: mdToSafeHtml(text) }} />;
}
