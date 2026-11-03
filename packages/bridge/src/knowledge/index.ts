import { watch, type FSWatcher } from "node:fs";
import { readdir, stat } from "node:fs/promises";
import { basename, join, resolve, sep } from "node:path";
import { homedir, platform } from "node:os";
import { chunkText, embed, embeddingPrefix, normalize, snippetOf, type KnowledgeStatus, type Policy, type Source, type SourceRef } from "@miai/shared";
import { appendAudit, auditFile } from "../audit.js";
import type { ConfigStore } from "../config.js";
import { extractText, isSupported, MAX_FILE_BYTES } from "./extract.js";
import { KnowledgeStore } from "./store.js";

/**
 * Lokal RAG over brugerens mapper (F15, ADR-0013): finder filer, holder indekset opdateret
 * (ved start, ved ændringer i mapperne og på forespørgsel) og søger i det.
 */

/** Højst så mange filer pr. indeks – beskytter mod at pege på en hel disk. */
export const MAX_FILES = 10_000;
/** Embeddings hentes i portioner af denne størrelse. */
const BATCH = 16;
/** Ventetid efter sidste filændring, før der indekseres igen. */
const WATCH_DEBOUNCE_MS = 3_000;
/** Uddrag med lavere lighed end dette gives ikke til modellen. */
const MIN_SCORE = 0.2;
/** Mapper, der aldrig gennemsøges. */
const SKIP_DIRS = new Set(["node_modules", ".git", "$RECYCLE.BIN", "System Volume Information"]);

/** "~/Dokumenter" → brugerens hjemmemappe + Dokumenter. */
export const expandHome = (p: string) => (p === "~" || p.startsWith("~/") || p.startsWith("~\\") ? join(homedir(), p.slice(1)) : p);

const caseInsensitive = () => platform() === "win32" || platform() === "darwin";
const norm = (p: string) => {
  const r = resolve(expandHome(p));
  return caseInsensitive() ? r.toLowerCase() : r;
};

/** Ligger mappen inden for en af politikkens tilladte mapper (eller er der ingen begrænsning)? */
export function folderAllowed(folder: string, policy: Policy): boolean {
  if (!policy.knowledge) return false;
  if (!policy.knowledgeFolders) return true;
  const f = norm(folder);
  return policy.knowledgeFolders.some((root) => {
    const r = norm(root);
    return f === r || f.startsWith(r.endsWith(sep) ? r : r + sep);
  });
}

async function walk(dir: string, out: string[], limit: number): Promise<boolean> {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return false;
  }
  for (const e of entries) {
    if (out.length >= limit) return true;
    // Skjulte filer og Office' låsefiler (~$Dokument.docx) springes over.
    if (e.name.startsWith(".") || e.name.startsWith("~$")) continue;
    const p = join(dir, e.name);
    if (e.isDirectory()) {
      if (SKIP_DIRS.has(e.name)) continue;
      if (await walk(p, out, limit)) return true;
    } else if (e.isFile() && isSupported(p)) out.push(p);
  }
  return out.length >= limit;
}

export class KnowledgeIndex {
  private store_: KnowledgeStore | undefined;
  /** Hvorfor indekset ikke kan bruges (fx ingen node:sqlite i denne runtime). */
  readonly unavailable: string | undefined;
  private status: KnowledgeStatus;
  private running: Promise<void> | undefined;
  private again = false;
  private watchers: FSWatcher[] = [];
  private timer: NodeJS.Timeout | undefined;
  private abort = new AbortController();
  private closed = false;

  constructor(
    private readonly store: ConfigStore,
    dbFile = join(store.dir, "knowledge.db"),
  ) {
    try {
      this.store_ = new KnowledgeStore(dbFile);
    } catch (err) {
      this.unavailable = `Vidensindekset kan ikke åbnes: ${(err as Error).message}. Kræver Node.js 22.13 eller nyere.`;
    }
    this.status = { state: "idle", done: 0, total: 0, documents: 0, chunks: 0, failed: [], skippedFolders: [], embeddingModel: "" };
  }

  get db(): KnowledgeStore {
    if (this.closed) throw new Error("Vidensindekset er lukket.");
    if (!this.store_) throw new Error(this.unavailable);
    return this.store_;
  }

  private get enabled(): boolean {
    if (this.unavailable) return false;
    const k = this.store.getSettings().knowledge;
    return k.enabled && this.store.getPolicy().knowledge && k.folders.length > 0;
  }

  getStatus(): KnowledgeStatus {
    const k = this.store.getSettings().knowledge;
    if (this.unavailable) return { ...this.status, state: k.enabled ? "error" : "disabled", error: this.unavailable, embeddingModel: k.embeddingModel };
    const s = this.db.stats();
    return {
      ...this.status,
      state: !k.enabled || !this.store.getPolicy().knowledge ? "disabled" : this.status.state,
      documents: s.documents,
      chunks: s.chunks,
      failed: this.db.errors(),
      embeddingModel: k.embeddingModel,
      lastIndexed: this.db.getMeta("lastIndexed"),
    };
  }

  /** Kaldes ved start og efter hver ændring af indstillingerne: opsætter overvågning og indekserer. */
  configure(): void {
    for (const w of this.watchers) w.close();
    this.watchers = [];
    if (!this.enabled) return;
    for (const folder of this.store.getSettings().knowledge.folders) {
      if (!folderAllowed(folder, this.store.getPolicy())) continue;
      try {
        const w = watch(expandHome(folder), { recursive: true }, (_event, name) => {
          if (name && isSupported(String(name))) this.schedule();
        });
        w.on("error", () => w.close());
        w.unref();
        this.watchers.push(w);
      } catch {
        // Mappen findes ikke, eller platformen kan ikke overvåge den – indeksering ved start og på knap.
      }
    }
    void this.reindex();
  }

  private schedule(): void {
    clearTimeout(this.timer);
    this.timer = setTimeout(() => void this.reindex(), WATCH_DEBOUNCE_MS);
    this.timer.unref();
  }

  /** Bringer indekset i overensstemmelse med mapperne. Kører én ad gangen; kald under en kørsel køres bagefter. */
  reindex(): Promise<void> {
    if (this.closed) return Promise.resolve();
    if (this.running) {
      this.again = true;
      return this.running;
    }
    this.running = this.run()
      .catch((err: Error) => {
        // En kørsel, der var i gang, da indekset blev lukket (fx ved nedlukning), er ikke en fejl.
        if (!this.closed) this.status = { ...this.status, state: "error", error: err.message };
      })
      .finally(() => {
        this.running = undefined;
        if (this.again && !this.closed) {
          this.again = false;
          void this.reindex();
        }
      });
    return this.running;
  }

  private async run(): Promise<void> {
    if (!this.enabled) {
      this.status = { ...this.status, state: "idle", done: 0, total: 0, error: undefined, skippedFolders: [] };
      return;
    }
    const settings = this.store.getSettings();
    const policy = this.store.getPolicy();
    const ep = this.store.activeEndpoint();
    const model = settings.knowledge.embeddingModel;
    const started = Date.now();

    // Vektorer fra forskellige modeller (eller endpoints) kan ikke sammenlignes: start forfra.
    const identity = `${ep.baseUrl}|${model}`;
    if (this.db.getMeta("embedding") !== identity) {
      this.db.clear();
      this.db.setMeta("embedding", identity);
    }

    const skippedFolders: KnowledgeStatus["skippedFolders"] = [];
    const files: string[] = [];
    let truncated = false;
    for (const folder of settings.knowledge.folders) {
      if (!folderAllowed(folder, policy)) {
        skippedFolders.push({ folder, reason: "policy" });
        continue;
      }
      const abs = resolve(expandHome(folder));
      const st = await stat(abs).catch(() => undefined);
      if (!st?.isDirectory()) {
        skippedFolders.push({ folder, reason: "missing" });
        continue;
      }
      truncated = (await walk(abs, files, MAX_FILES)) || truncated;
    }

    const known = new Map(this.db.docs().map((d) => [d.path, d]));
    const seen = new Set(files);
    let removed = 0;
    for (const path of known.keys()) {
      if (!seen.has(path)) {
        this.db.removeDoc(path);
        removed++;
      }
    }

    const todo: { path: string; mtime: number; size: number }[] = [];
    for (const path of files) {
      const st = await stat(path).catch(() => undefined);
      if (!st) continue;
      const k = known.get(path);
      if (k && k.mtime === st.mtimeMs && k.size === st.size) continue;
      todo.push({ path, mtime: st.mtimeMs, size: st.size });
    }

    this.status = { ...this.status, state: "indexing", done: 0, total: todo.length, skippedFolders, truncated, error: undefined };
    let indexed = 0;
    try {
      for (const f of todo) {
        await this.indexFile(f, ep, model);
        indexed++;
        this.status = { ...this.status, done: indexed };
      }
      this.db.setMeta("lastIndexed", new Date().toISOString());
      this.status = { ...this.status, state: "idle" };
    } catch (err) {
      // Fejl fra endpointen (fx ingen embeddings) stopper kørslen; filfejl registreres pr. fil.
      this.status = { ...this.status, state: "error", error: (err as Error).message };
    }
    if (settings.auditLog && (indexed || removed)) {
      await appendAudit(auditFile(this.store.dir), {
        ts: new Date().toISOString(),
        kind: "index",
        host: new URL(ep.baseUrl).host,
        model,
        status: this.status.state === "error" ? "error" : "ok",
        durationMs: Date.now() - started,
        count: indexed + removed,
      });
    }
  }

  private async indexFile(f: { path: string; mtime: number; size: number }, ep: ReturnType<ConfigStore["activeEndpoint"]>, model: string) {
    const title = basename(f.path);
    if (f.size > MAX_FILE_BYTES) {
      this.db.putDoc({ ...f, title }, [], "Filen er for stor (over 30 MB).");
      return;
    }
    let text: string;
    try {
      text = await extractText(f.path);
    } catch (err) {
      this.db.putDoc({ ...f, title }, [], `Kunne ikke læse filen: ${(err as Error).message}`.slice(0, 500));
      return;
    }
    const chunks = chunkText(text);
    const vectors: Float32Array[] = [];
    const prefix = embeddingPrefix(model, "document");
    for (let i = 0; i < chunks.length; i += BATCH) {
      // Filnavnet står foran hvert uddrag – det hjælper søgningen (fx "Rejsepolitik.docx").
      const batch = chunks.slice(i, i + BATCH).map((c) => `${prefix}${title}\n${c}`);
      const out = await embed(ep, model, batch, this.abort.signal);
      vectors.push(...out.map(normalize));
    }
    this.db.putDoc(
      { ...f, title },
      chunks.map((text, i) => ({ text, vector: vectors[i] })),
    );
  }

  /**
   * Finder de bedste uddrag til et spørgsmål. Returnerer både det, modellen får (Source), og
   * det, brugeren ser (SourceRef).
   */
  async search(query: string, k = this.store.getSettings().knowledge.topK, signal?: AbortSignal): Promise<{ sources: Source[]; refs: SourceRef[] }> {
    const model = this.store.getSettings().knowledge.embeddingModel;
    const [q] = await embed(this.store.activeEndpoint(), model, [embeddingPrefix(model, "query") + query], signal);
    const hits = this.db.search(normalize(q), k, MIN_SCORE);
    return {
      sources: hits.map((h, i) => ({ n: i + 1, title: h.title, path: h.path, text: h.text })),
      refs: hits.map((h, i) => ({
        n: i + 1,
        title: h.title,
        path: h.path,
        chunkId: h.chunkId,
        score: Math.round(h.score * 100) / 100,
        snippet: snippetOf(h.text),
      })),
    };
  }

  close(): void {
    this.closed = true;
    this.abort.abort();
    clearTimeout(this.timer);
    for (const w of this.watchers) w.close();
    this.store_?.close();
  }
}
