/**
 * Samtalen huskes pr. dokument (#26): beskederne gemmes i task panes localStorage under en nøgle
 * afledt af dokumentets adresse, så samtalen er der igen, når task pane åbnes på samme dokument.
 *
 * Kun på maskinen: intet sendes nogen steder hen. Der gemmes tekst – ikke agentens planer med
 * dokumentkontekst – og mængden er begrænset (beskeder pr. dokument, tegn og antal dokumenter).
 * Dokumenter uden adresse (ikke gemt endnu) huskes ikke.
 */
export interface StoredMessage {
  role: "user" | "assistant";
  content: string;
  model?: string;
  /** Svaret var en plan fra agenten; `log` er udfaldet, som modellen skal kende i opfølgninger. */
  planned?: boolean;
  log?: string;
}

type Store = Pick<Storage, "getItem" | "setItem" | "removeItem">;

const PREFIX = "miai.chat.v1:";
const INDEX = "miai.chat.v1.index";
export const MAX_STORED_MESSAGES = 40;
export const MAX_STORED_CHARS = 60_000;
export const MAX_STORED_DOCUMENTS = 20;

/** Kort, stabil nøgle for dokumentets adresse (FNV-1a) – stien selv gemmes ikke som nøgle. */
export function documentKey(url: string | null | undefined): string | null {
  if (!url) return null;
  let h = 0x811c9dc5;
  for (let i = 0; i < url.length; i++) {
    h ^= url.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16).padStart(8, "0") + url.length.toString(16);
}

const readIndex = (store: Store): string[] => {
  try {
    const v = JSON.parse(store.getItem(INDEX) ?? "[]") as unknown;
    return Array.isArray(v) ? v.filter((x): x is string => typeof x === "string") : [];
  } catch {
    return [];
  }
};

export function loadChat(key: string | null, store: Store = localStorage): StoredMessage[] {
  if (!key) return [];
  try {
    const v = JSON.parse(store.getItem(PREFIX + key) ?? "[]") as unknown;
    if (!Array.isArray(v)) return [];
    return v.filter((m): m is StoredMessage => !!m && typeof m === "object" && (m.role === "user" || m.role === "assistant") && typeof m.content === "string");
  } catch {
    return [];
  }
}

/** Gemmer de seneste beskeder for dokumentet; de ældste dokumenter ryger ud, når der er for mange. */
export function saveChat(key: string | null, messages: StoredMessage[], store: Store = localStorage): void {
  if (!key) return;
  try {
    if (!messages.length) return clearChat(key, store);
    // De nyeste beskeder inden for grænserne (bagfra, så en lang samtale beholder slutningen).
    const kept: StoredMessage[] = [];
    let chars = 0;
    for (const m of [...messages].reverse()) {
      chars += m.content.length + (m.log?.length ?? 0);
      if (kept.length >= MAX_STORED_MESSAGES || (kept.length && chars > MAX_STORED_CHARS)) break;
      kept.unshift(m);
    }
    store.setItem(PREFIX + key, JSON.stringify(kept));
    const index = [key, ...readIndex(store).filter((k) => k !== key)];
    for (const old of index.slice(MAX_STORED_DOCUMENTS)) store.removeItem(PREFIX + old);
    store.setItem(INDEX, JSON.stringify(index.slice(0, MAX_STORED_DOCUMENTS)));
  } catch {
    // Fuldt eller spærret lager: samtalen virker stadig, den huskes bare ikke.
  }
}

export function clearChat(key: string | null, store: Store = localStorage): void {
  if (!key) return;
  try {
    store.removeItem(PREFIX + key);
    store.setItem(INDEX, JSON.stringify(readIndex(store).filter((k) => k !== key)));
  } catch {
    // Se saveChat.
  }
}
