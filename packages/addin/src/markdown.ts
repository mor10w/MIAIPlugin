import DOMPurify from "dompurify";
import { marked } from "marked";

/**
 * Modeloutput er utroværdigt (dokumentindhold kan indeholde prompt injection),
 * så al HTML renses, før den vises i task pane eller indsættes i Office.
 */
export function mdToSafeHtml(md: string): string {
  const html = marked.parse(md, { async: false, gfm: true, breaks: false }) as string;
  return DOMPurify.sanitize(html, { USE_PROFILES: { html: true }, FORBID_TAGS: ["style", "form", "input", "img", "iframe"], FORBID_ATTR: ["style"] });
}

/** Markdown → ren tekst (til PowerPoint og Excel, som ikke tager HTML). */
export function mdToPlain(md: string): string {
  return md
    .replace(/```[a-z]*\n?/gi, "")
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/^\s*[-*+]\s+/gm, "• ")
    .replace(/\*\*(.+?)\*\*/g, "$1")
    .replace(/(^|\W)[*_](.+?)[*_](?=\W|$)/g, "$1$2")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .trim();
}
