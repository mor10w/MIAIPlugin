import { readFile } from "node:fs/promises";
import { extname } from "node:path";
import JSZip from "jszip";

/**
 * Tekst fra brugerens filer til lokal RAG (F15, ADR-0013). Kun ren JavaScript (ingen native
 * moduler), så bridgen og MIAI App kan pakkes uden byggetrin hos brugeren.
 *
 * Understøttet: Word (.docx), PDF, Excel (.xlsx), PowerPoint (.pptx), Markdown og tekst.
 */
export const SUPPORTED_EXTENSIONS = [".docx", ".pdf", ".xlsx", ".pptx", ".md", ".markdown", ".txt", ".csv"] as const;

export const isSupported = (file: string) => (SUPPORTED_EXTENSIONS as readonly string[]).includes(extname(file).toLowerCase());

/** Filer over denne størrelse springes over (store PDF'er med billeder giver sjældent meget tekst). */
export const MAX_FILE_BYTES = 30 * 1024 * 1024;

/** Maks. tekst pr. fil, der indekseres (≈ 400 sider). */
export const MAX_TEXT_CHARS = 1_000_000;

export async function extractText(file: string): Promise<string> {
  const ext = extname(file).toLowerCase();
  const buf = await readFile(file);
  let text: string;
  switch (ext) {
    case ".docx":
      text = await docx(buf);
      break;
    case ".pdf":
      text = await pdf(buf);
      break;
    case ".xlsx":
      text = await xlsx(buf);
      break;
    case ".pptx":
      text = await pptx(buf);
      break;
    default:
      text = buf.toString("utf8").replace(/^\uFEFF/, "");
  }
  return text.slice(0, MAX_TEXT_CHARS);
}

async function docx(buf: Buffer): Promise<string> {
  const mammoth = await import("mammoth");
  const r = await (mammoth.default ?? mammoth).extractRawText({ buffer: buf });
  return r.value;
}

async function pdf(buf: Buffer): Promise<string> {
  const { extractText: pdfText, getDocumentProxy } = await import("unpdf");
  const doc = await getDocumentProxy(new Uint8Array(buf));
  const { text } = await pdfText(doc, { mergePages: false });
  // Sidetal i teksten, så uddrag og kildehenvisning kan pege på en side.
  return (text as string[]).map((t, i) => `[Side ${i + 1}]\n${t}`).join("\n\n");
}

const decode = (s: string) =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, "&");

/** Alle <tag>…</tag>-tekster i en XML-streng (fx <a:t> i PowerPoint, <t> i Excel). */
const texts = (xml: string, tag: string) => [...xml.matchAll(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, "g"))].map((m) => decode(m[1]));

const byNumber = (prefix: string) => (a: string, b: string) =>
  Number(a.slice(prefix.length).match(/\d+/)?.[0]) - Number(b.slice(prefix.length).match(/\d+/)?.[0]);

async function pptx(buf: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  const slides = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort(byNumber("ppt/slides/slide"));
  const out: string[] = [];
  for (const [i, f] of slides.entries()) {
    const xml = await zip.file(f)!.async("string");
    // Ét afsnit (<a:p>) pr. linje.
    const paras = [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) => texts(p[1], "a:t").join("")).filter((t) => t.trim());
    out.push(`[Slide ${i + 1}]\n${paras.join("\n")}`);
  }
  return out.join("\n\n");
}

/** Excel som tekst: én linje pr. række med tabulator mellem cellerne, ark for ark. */
async function xlsx(buf: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  const shared = zip.file("xl/sharedStrings.xml");
  const strings = shared ? [...(await shared.async("string")).matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => texts(m[1], "t").join("")) : [];
  const wb = (await zip.file("xl/workbook.xml")?.async("string")) ?? "";
  const names = [...wb.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((m) => decode(m[1]));
  const sheets = Object.keys(zip.files)
    .filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f))
    .sort(byNumber("xl/worksheets/sheet"));
  const out: string[] = [];
  for (const [i, f] of sheets.entries()) {
    const xml = await zip.file(f)!.async("string");
    const rows: string[] = [];
    for (const row of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
      const cells = [...row[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)].map((c) => {
        const attrs = c[1];
        const body = c[2] ?? "";
        const v = body.match(/<v>([\s\S]*?)<\/v>/)?.[1];
        if (/\bt="s"/.test(attrs) && v !== undefined) return strings[Number(v)] ?? "";
        if (/\bt="inlineStr"/.test(attrs)) return texts(body, "t").join("");
        return v !== undefined ? decode(v) : "";
      });
      if (cells.some((x) => x.trim())) rows.push(cells.join("\t"));
    }
    if (rows.length) out.push(`[Ark: ${names[i] ?? `Ark${i + 1}`}]\n${rows.join("\n")}`);
  }
  return out.join("\n\n");
}
