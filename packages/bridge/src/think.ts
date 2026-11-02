/**
 * Fjerner <think>…</think>-blokke fra en tekststrøm, også når tags er splittet over flere chunks.
 * Mange reasoning-modeller (Qwen3, DeepSeek-R1 m.fl.) sender deres tankegang i content-feltet.
 */
export class ThinkFilter {
  private buf = "";
  private inThink = false;
  private static OPEN = "<think>";
  private static CLOSE = "</think>";

  push(chunk: string): string {
    this.buf += chunk;
    let out = "";
    for (;;) {
      if (this.inThink) {
        const end = this.buf.indexOf(ThinkFilter.CLOSE);
        if (end < 0) {
          // behold en hale, der kan være starten på </think>
          this.buf = this.buf.slice(Math.max(0, this.buf.length - (ThinkFilter.CLOSE.length - 1)));
          return out;
        }
        this.buf = this.buf.slice(end + ThinkFilter.CLOSE.length);
        this.inThink = false;
        continue;
      }
      const start = this.buf.indexOf(ThinkFilter.OPEN);
      if (start >= 0) {
        out += this.buf.slice(0, start);
        this.buf = this.buf.slice(start + ThinkFilter.OPEN.length);
        this.inThink = true;
        continue;
      }
      // Udsend alt undtagen en mulig delvis "<think" i slutningen
      const keep = partialSuffix(this.buf, ThinkFilter.OPEN);
      out += this.buf.slice(0, this.buf.length - keep);
      this.buf = this.buf.slice(this.buf.length - keep);
      return out;
    }
  }

  flush(): string {
    const rest = this.inThink ? "" : this.buf;
    this.buf = "";
    return rest;
  }
}

/** Længden af det længste suffiks af `s`, der er et præfiks af `tag`. */
function partialSuffix(s: string, tag: string): number {
  for (let n = Math.min(s.length, tag.length - 1); n > 0; n--) {
    if (tag.startsWith(s.slice(-n))) return n;
  }
  return 0;
}

/** Fjerner foranstillede blanke linjer, som modeller ofte sender efter </think>. */
export function trimLeading(s: string): string {
  return s.replace(/^\s+/, "");
}
