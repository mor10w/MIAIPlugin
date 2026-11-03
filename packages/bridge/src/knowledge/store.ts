import { mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import type { DatabaseSync } from "node:sqlite";
import { topK } from "@miai/shared";

/**
 * Vidensindekset (F15, ADR-0013) i én SQLite-fil ved siden af config.json, via Nodes indbyggede
 * `node:sqlite` (ingen native moduler). Embeddings gemmes normaliserede som Float32-blobs og
 * søges med brute force i hukommelsen.
 */

export interface StoredDoc {
  id: number;
  path: string;
  mtime: number;
  size: number;
  chunks: number;
  error: string | null;
}

export interface ChunkHit {
  chunkId: number;
  score: number;
  text: string;
  path: string;
  title: string;
}

export class KnowledgeStore {
  private db: DatabaseSync;
  /** Alle vektorer i hukommelsen; nulstilles ved ændringer og læses igen ved næste søgning. */
  private cache: { ids: number[]; vectors: Float32Array[] } | undefined;

  constructor(readonly file: string) {
    // Indlæses først her (ikke ved import), så en runtime uden node:sqlite kun slår vidensindekset
    // fra – resten af bridgen virker (se KnowledgeIndex).
    const { DatabaseSync: Db } = createRequire(import.meta.url)("node:sqlite") as typeof import("node:sqlite");
    if (file !== ":memory:") mkdirSync(dirname(file), { recursive: true });
    this.db = new Db(file);
    this.db.exec(`
      PRAGMA journal_mode = WAL;
      PRAGMA foreign_keys = ON;
      CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY,
        path TEXT NOT NULL UNIQUE,
        title TEXT NOT NULL,
        mtime REAL NOT NULL,
        size INTEGER NOT NULL,
        chunks INTEGER NOT NULL DEFAULT 0,
        error TEXT,
        indexed_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS chunks (
        id INTEGER PRIMARY KEY,
        doc_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
        ord INTEGER NOT NULL,
        text TEXT NOT NULL,
        embedding BLOB NOT NULL
      );
      CREATE INDEX IF NOT EXISTS chunks_doc ON chunks(doc_id);
    `);
  }

  getMeta(key: string): string | undefined {
    return (this.db.prepare("SELECT value FROM meta WHERE key = ?").get(key) as { value: string } | undefined)?.value;
  }

  setMeta(key: string, value: string): void {
    this.db.prepare("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(key, value);
  }

  docs(): StoredDoc[] {
    return this.db.prepare("SELECT id, path, mtime, size, chunks, error FROM documents").all() as unknown as StoredDoc[];
  }

  /** Erstatter en fils uddrag (i én transaktion, så en afbrudt indeksering ikke efterlader halve filer). */
  putDoc(doc: { path: string; title: string; mtime: number; size: number }, chunks: { text: string; vector: Float32Array }[], error?: string): void {
    this.db.exec("BEGIN");
    try {
      this.db.prepare("DELETE FROM documents WHERE path = ?").run(doc.path);
      const r = this.db
        .prepare("INSERT INTO documents (path, title, mtime, size, chunks, error, indexed_at) VALUES (?, ?, ?, ?, ?, ?, ?)")
        .run(doc.path, doc.title, doc.mtime, doc.size, chunks.length, error ?? null, new Date().toISOString());
      const insert = this.db.prepare("INSERT INTO chunks (doc_id, ord, text, embedding) VALUES (?, ?, ?, ?)");
      chunks.forEach((c, i) => insert.run(r.lastInsertRowid, i, c.text, new Uint8Array(c.vector.buffer, c.vector.byteOffset, c.vector.byteLength)));
      this.db.exec("COMMIT");
    } catch (err) {
      this.db.exec("ROLLBACK");
      throw err;
    }
    this.cache = undefined;
  }

  removeDoc(path: string): void {
    this.db.prepare("DELETE FROM documents WHERE path = ?").run(path);
    this.cache = undefined;
  }

  /** Tømmer indekset (fx når embedding-modellen skiftes – vektorer fra to modeller kan ikke sammenlignes). */
  clear(): void {
    this.db.exec("DELETE FROM documents; DELETE FROM meta;");
    this.cache = undefined;
  }

  stats(): { documents: number; chunks: number; errors: number } {
    const r = this.db
      .prepare("SELECT COUNT(*) AS documents, COALESCE(SUM(chunks), 0) AS chunks, COALESCE(SUM(error IS NOT NULL), 0) AS errors FROM documents")
      .get() as { documents: number; chunks: number; errors: number };
    return { documents: Number(r.documents), chunks: Number(r.chunks), errors: Number(r.errors) };
  }

  errors(limit = 20): { path: string; error: string }[] {
    return this.db.prepare("SELECT path, error FROM documents WHERE error IS NOT NULL ORDER BY indexed_at DESC LIMIT ?").all(limit) as unknown as {
      path: string;
      error: string;
    }[];
  }

  private load() {
    if (this.cache) return this.cache;
    const rows = this.db.prepare("SELECT id, embedding FROM chunks").all() as { id: number; embedding: Uint8Array }[];
    // Kopiér til en justeret buffer (SQLite-blobs er ikke nødvendigvis 4-byte-justerede).
    this.cache = {
      ids: rows.map((r) => Number(r.id)),
      vectors: rows.map((r) => new Float32Array(new Uint8Array(r.embedding).buffer)),
    };
    return this.cache;
  }

  search(query: Float32Array, k: number, minScore: number): ChunkHit[] {
    const { ids, vectors } = this.load();
    const best = topK(query, vectors, k, minScore);
    const get = this.db.prepare("SELECT c.text, d.path, d.title FROM chunks c JOIN documents d ON d.id = c.doc_id WHERE c.id = ?");
    return best.map(({ index, score }) => {
      const row = get.get(ids[index]) as { text: string; path: string; title: string };
      return { chunkId: ids[index], score, ...row };
    });
  }

  chunk(id: number): { text: string; path: string; title: string } | undefined {
    return this.db.prepare("SELECT c.text, d.path, d.title FROM chunks c JOIN documents d ON d.id = c.doc_id WHERE c.id = ?").get(id) as
      { text: string; path: string; title: string } | undefined;
  }

  close(): void {
    this.db.close();
  }
}
