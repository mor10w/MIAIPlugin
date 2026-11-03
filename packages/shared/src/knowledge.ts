/**
 * Rene hjælpefunktioner til lokal RAG (F15, ADR-0013): opdeling af tekst i uddrag og
 * vektorsøgning. Indeksering og filadgang ligger i bridgen.
 */

/** Uddragets mål- og maks.-længde i tegn (~250–400 tokens) og overlap mellem nabo-uddrag. */
export const CHUNK_TARGET = 1_200;
export const CHUNK_MAX = 1_800;
export const CHUNK_OVERLAP = 200;

/**
 * Deler tekst i uddrag ved afsnitsgrænser (tomme linjer), så et uddrag så vidt muligt er
 * hele afsnit. Lange afsnit deles ved sætningsgrænser; nabo-uddrag overlapper lidt, så en
 * pointe, der krydser grænsen, findes fra begge sider.
 */
export function chunkText(text: string, target = CHUNK_TARGET, max = CHUNK_MAX, overlap = CHUNK_OVERLAP): string[] {
  const clean = text
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  if (!clean) return [];
  // Afsnit; for lange afsnit deles i sætninger, og helt uden sætningsgrænser i faste stykker.
  const pieces: string[] = [];
  for (const para of clean.split(/\n\s*\n/)) {
    if (para.length <= max) {
      pieces.push(para);
      continue;
    }
    const sentences = para.match(/[^.!?\n]+(?:[.!?]+|\n|$)\s*/g) ?? [para];
    let cur = "";
    for (const s of sentences) {
      if (s.length > max) {
        if (cur) pieces.push(cur.trim());
        cur = "";
        for (let i = 0; i < s.length; i += target) pieces.push(s.slice(i, i + target).trim());
        continue;
      }
      if ((cur + s).length > target && cur) {
        pieces.push(cur.trim());
        cur = "";
      }
      cur += s;
    }
    if (cur.trim()) pieces.push(cur.trim());
  }

  const chunks: string[] = [];
  let cur = "";
  for (const p of pieces) {
    if (cur && cur.length + p.length + 2 > target) {
      chunks.push(cur);
      // Overlap: slutningen af forrige uddrag (fra en ordgrænse) indleder det næste.
      const tail = cur.slice(-overlap);
      const cut = tail.indexOf(" ");
      cur = overlap > 0 && cut >= 0 ? `${tail.slice(cut + 1)}\n\n${p}` : p;
    } else {
      cur = cur ? `${cur}\n\n${p}` : p;
    }
  }
  if (cur) chunks.push(cur);
  return chunks;
}

/** Normaliserer en vektor til længde 1, så cosinus-lighed blot er prikproduktet. */
export function normalize(v: ArrayLike<number>): Float32Array {
  let sum = 0;
  for (let i = 0; i < v.length; i++) sum += v[i] * v[i];
  const n = Math.sqrt(sum) || 1;
  const out = new Float32Array(v.length);
  for (let i = 0; i < v.length; i++) out[i] = v[i] / n;
  return out;
}

export function dot(a: Float32Array, b: Float32Array): number {
  let s = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) s += a[i] * b[i];
  return s;
}

/**
 * De `k` bedste (normaliserede) vektorer for forespørgslen, over en minimumsgrænse.
 * Brute force: til personlige mapper (op til ~100.000 uddrag) tager det få millisekunder
 * og kræver ingen native udvidelse (ADR-0013).
 */
export function topK(query: Float32Array, vectors: ArrayLike<Float32Array>, k: number, minScore = 0): { index: number; score: number }[] {
  const best: { index: number; score: number }[] = [];
  for (let i = 0; i < vectors.length; i++) {
    const score = dot(query, vectors[i]);
    if (score < minScore) continue;
    if (best.length < k) {
      best.push({ index: i, score });
      best.sort((a, b) => b.score - a.score);
    } else if (score > best[k - 1].score) {
      best[k - 1] = { index: i, score };
      best.sort((a, b) => b.score - a.score);
    }
  }
  return best;
}

/** Kort uddrag til visning: første ~240 tegn, afsluttet ved et ord. */
export function snippetOf(text: string, max = 240): string {
  const t = text.replace(/\s+/g, " ").trim();
  if (t.length <= max) return t;
  const cut = t.lastIndexOf(" ", max);
  return `${t.slice(0, cut > max / 2 ? cut : max)}…`;
}
