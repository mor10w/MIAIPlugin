import { useEffect, useRef } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { mdToSafeHtml } from "../markdown";
import { ui } from "../theme";
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
    "& code": {
      fontFamily: tokens.fontFamilyMonospace,
      background: tokens.colorNeutralBackground3,
      padding: `0 ${tokens.spacingHorizontalXXS}`,
      borderRadius: tokens.borderRadiusSmall,
    },
    "& pre": { background: tokens.colorNeutralBackground3, padding: tokens.spacingHorizontalS, overflowX: "auto" },
    // Figurer står altid på hvidt "papir" – Mermaid/SVG er tegnet til lys baggrund og skal kunne læses i mørkt tema.
    "& .figure": {
      margin: `0 0 ${tokens.spacingVerticalS}`,
      overflowX: "auto",
      padding: tokens.spacingHorizontalS,
      borderRadius: tokens.borderRadiusMedium,
      backgroundColor: ui.paper,
      color: ui.paperInk,
      border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
    },
    "& a": { color: tokens.colorBrandForegroundLink },
    "& > :last-child": { marginBottom: 0 },
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
export function Markdown({
  text,
  onDiagramError,
}: {
  text: string;
  /** Et mermaid-diagram kunne ikke tegnes (heller ikke efter reparation) – med mermaids fejl. */
  onDiagramError?: (error: string | null) => void;
}) {
  const s = useStyles();
  const ref = useRef<HTMLDivElement>(null);

  // Ref: kalderen giver en ny funktion ved hver rendering, og effekten skal kun køre, når teksten ændres.
  const onErrorRef = useRef(onDiagramError);
  onErrorRef.current = onDiagramError;

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
      let failed: string | null = null;
      for (const block of Array.from(root.querySelectorAll("code.language-mermaid"))) {
        try {
          const svg = await mermaidToSvg(block.textContent ?? "");
          if (cancelled) return;
          const wrap = document.createElement("div");
          wrap.className = "figure";
          // SVG'en kommer fra mermaid.js (securityLevel: strict), ikke fra modeloutput.
          wrap.innerHTML = svg;
          replaceBlock(block, wrap);
        } catch (err) {
          // Ugyldigt diagram: vis kodeblokken som den er, og meld fejlen (chatten kan bede modellen rette den).
          failed ??= String((err as Error)?.message ?? err).slice(0, 600);
        }
      }
      if (!cancelled) onErrorRef.current?.(failed);
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

  // eslint-disable-next-line no-restricted-syntax -- mdToSafeHtml renser modellens Markdown med DOMPurify.
  return <div ref={ref} className={s.md} dangerouslySetInnerHTML={{ __html: mdToSafeHtml(text) }} />;
}
