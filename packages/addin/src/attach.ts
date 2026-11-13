import { MAX_ATTACHMENT_TEXT, MAX_IMAGE_CHARS, type Attachment } from "@miai/shared";

/**
 * Filer, brugeren vedhæfter en chatbesked (#27): billeder skaleres ned og sendes som data-URL til
 * modeller med vision; tekstfiler læses som tekst. Alt sker i task pane – intet uploades.
 */

/** Længste side på et vedhæftet billede. Større billeder skaleres ned (modellerne gør det alligevel). */
export const MAX_IMAGE_SIDE = 1568;
const TEXT_EXTENSIONS = /\.(txt|md|markdown|csv|tsv|json|log|xml|ya?ml|html?|sql|py|js|ts)$/i;
/** Dokumentfiler, som kun bridgen kan trække tekst ud af (samme udtræk som videnbasen). */
export const DOCUMENT_EXTENSIONS = /\.(pdf|docx|xlsx|pptx)$/i;
export const IMAGE_TYPES = ["image/png", "image/jpeg", "image/webp", "image/gif"];
export const ACCEPT = [...IMAGE_TYPES, ".txt,.md,.csv,.tsv,.json,.log,.xml,.yaml,.yml,.html,.sql", ".pdf,.docx,.xlsx,.pptx"].join(",");

export type AttachError = "too-large" | "unsupported" | "needs-bridge" | "unreadable";
export class AttachmentError extends Error {
  constructor(
    readonly reason: AttachError,
    readonly file: string,
  ) {
    super(reason);
  }
}

/** Ny størrelse, så længste side højst er MAX_IMAGE_SIDE (forholdet bevares). */
export function fitImage(width: number, height: number, max = MAX_IMAGE_SIDE): { width: number; height: number } {
  const scale = Math.min(1, max / Math.max(width, height, 1));
  return { width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)) };
}

/** Afkorter tekst til grænsen og siger det i teksten, så modellen ved, at resten mangler. */
export function clipText(text: string, max = MAX_ATTACHMENT_TEXT): string {
  const note = "\n[… the rest of the file was left out]";
  return text.length > max ? text.slice(0, max - note.length) + note : text;
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("image"));
    img.src = url;
  });
}

async function imageAttachment(file: File): Promise<Attachment> {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url).catch(() => {
      throw new AttachmentError("unreadable", file.name);
    });
    const { width, height } = fitImage(img.naturalWidth, img.naturalHeight);
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new AttachmentError("unreadable", file.name);
    // Hvid baggrund: gennemsigtige skærmbilleder bliver ellers sorte som JPEG.
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, width, height);
    ctx.drawImage(img, 0, 0, width, height);
    // PNG holder tekst i skærmbilleder skarp; fylder den for meget, bruges JPEG.
    let dataUrl = canvas.toDataURL("image/png");
    if (dataUrl.length > MAX_IMAGE_CHARS / 2) dataUrl = canvas.toDataURL("image/jpeg", 0.85);
    if (dataUrl.length > MAX_IMAGE_CHARS) throw new AttachmentError("too-large", file.name);
    return { kind: "image", name: file.name || "billede.png", dataUrl };
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * Gør en fil til en vedhæftning. `extract` er bridgens udtræk af tekst fra PDF og Office-filer;
 * uden bridge kan de ikke vedhæftes.
 */
export async function toAttachment(file: File, extract?: (file: File) => Promise<string>): Promise<Attachment> {
  if (IMAGE_TYPES.includes(file.type)) return imageAttachment(file);
  if (DOCUMENT_EXTENSIONS.test(file.name)) {
    if (!extract) throw new AttachmentError("needs-bridge", file.name);
    if (file.size > 20 * 1024 * 1024) throw new AttachmentError("too-large", file.name);
    return { kind: "text", name: file.name, text: clipText(await extract(file)) };
  }
  if (file.type.startsWith("text/") || file.type === "application/json" || TEXT_EXTENSIONS.test(file.name)) {
    if (file.size > 5 * 1024 * 1024) throw new AttachmentError("too-large", file.name);
    return { kind: "text", name: file.name, text: clipText(await file.text()) };
  }
  throw new AttachmentError("unsupported", file.name);
}
