/**
 * Cache for regnearksfunktionen =MIAI.ASK() (#80). Excel kalder funktionen igen ved F9 og muligvis
 * ved åbning af projektmappen; uden cache ville 1.000 celler give 1.000 modelkald hver gang.
 *
 * Cachen overlever genstart (localStorage), har et loft (ældste ud først) og gemmer kun en hash af
 * forespørgslen sammen med svaret – aldrig prompt eller input i klartekst.
 */

export const CELL_CACHE_KEY = "miai.cellcache.v1";
export const MAX_CACHE_ENTRIES = 5_000;
/** Samlet størrelse af de gemte svar i tegn (localStorage har typisk ca. 5 MB pr. adresse). */
export const MAX_CACHE_CHARS = 1_000_000;

export interface CacheStorage {
  get(key: string): string | null;
  set(key: string, value: string): void;
  remove(key: string): void;
}

/** localStorage, som tåler at være slået fra (privat tilstand, fuld kvote). */
export const browserStorage: CacheStorage = {
  get: (k) => {
    try {
      return localStorage.getItem(k);
    } catch {
      return null;
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, v);
    } catch {
      /* fuld eller slået fra: cachen lever så kun i hukommelsen */
    }
  },
  remove: (k) => {
    try {
      localStorage.removeItem(k);
    } catch {
      /* ignorér */
    }
  },
};

/** Det, der afgør svaret. Ændres noget af det, er det en anden forespørgsel. */
export interface CellQuery {
  endpoint: string;
  model: string;
  optimized: boolean;
  prompt: string;
  input: string;
}

const hex = (bytes: Uint8Array) => Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");

/** Reserve, hvis WebCrypto mangler: to uafhængige 53-bit-hashes (cyrb53) sat sammen. */
function cyrb(str: string, seed: number): string {
  let h1 = 0xdeadbeef ^ seed;
  let h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).padStart(14, "0");
}

/** Nøglen for en forespørgsel: de første 128 bit af SHA-256 over alle felter. */
export async function queryKey(q: CellQuery): Promise<string> {
  // Længderne med, så felterne ikke kan flyde sammen ("a" + "bc" ≠ "ab" + "c").
  const text = ["v1", q.endpoint, q.model, q.optimized ? "1" : "0", q.prompt.length, q.prompt, q.input].join("\u0000");
  const subtle = globalThis.crypto?.subtle;
  if (subtle) return hex(new Uint8Array(await subtle.digest("SHA-256", new TextEncoder().encode(text))).slice(0, 16));
  return cyrb(text, 1) + cyrb(text, 2);
}

export class CellCache {
  /** Map bevarer indsættelsesrækkefølgen: det først indsatte er det ældst brugte. */
  private entries = new Map<string, string>();
  private chars = 0;
  private timer: ReturnType<typeof setTimeout> | undefined;

  constructor(
    private readonly storage: CacheStorage = browserStorage,
    private readonly limits = { entries: MAX_CACHE_ENTRIES, chars: MAX_CACHE_CHARS },
  ) {
    try {
      const saved = JSON.parse(storage.get(CELL_CACHE_KEY) ?? "[]") as unknown;
      if (Array.isArray(saved)) for (const e of saved) if (Array.isArray(e) && typeof e[0] === "string" && typeof e[1] === "string") this.put(e[0], e[1]);
    } catch {
      /* ødelagt indhold: start forfra */
    }
  }

  get size(): number {
    return this.entries.size;
  }

  get(key: string): string | undefined {
    const v = this.entries.get(key);
    if (v === undefined) return undefined;
    // Flyt til enden: senest brugt.
    this.entries.delete(key);
    this.entries.set(key, v);
    return v;
  }

  set(key: string, value: string): void {
    this.put(key, value);
    // Saml skrivningerne: en kørsel over mange celler skal ikke skrive hele cachen for hver celle.
    this.timer ??= setTimeout(() => this.flush(), 500);
  }

  /** Skriver cachen til lageret nu (også ved pagehide). */
  flush(): void {
    if (this.timer) clearTimeout(this.timer);
    this.timer = undefined;
    this.storage.set(CELL_CACHE_KEY, JSON.stringify([...this.entries]));
  }

  clear(): void {
    if (this.timer) clearTimeout(this.timer);
    this.timer = undefined;
    this.entries.clear();
    this.chars = 0;
    this.storage.remove(CELL_CACHE_KEY);
  }

  private put(key: string, value: string): void {
    const old = this.entries.get(key);
    if (old !== undefined) {
      this.chars -= old.length;
      this.entries.delete(key);
    }
    // Et enkelt svar større end hele loftet gemmes ikke.
    if (value.length > this.limits.chars) return;
    this.entries.set(key, value);
    this.chars += value.length;
    while (this.entries.size > this.limits.entries || this.chars > this.limits.chars) {
      const [oldest, v] = this.entries.entries().next().value as [string, string];
      this.entries.delete(oldest);
      this.chars -= v.length;
    }
  }
}
