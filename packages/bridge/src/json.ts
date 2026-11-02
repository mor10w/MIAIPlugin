/**
 * Udtrækker et JSON-objekt fra modeltekst. Tåler code fences og tekst før/efter,
 * fordi ikke alle endpoints understøtter struktureret output.
 */
export function extractJson(text: string): unknown {
  const t = text.trim();
  try {
    return JSON.parse(t);
  } catch {
    /* prøv videre */
  }
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) {
    try {
      return JSON.parse(fence[1].trim());
    } catch {
      /* prøv videre */
    }
  }
  const start = t.search(/[[{]/);
  if (start >= 0) {
    const open = t[start];
    const close = open === "{" ? "}" : "]";
    let depth = 0;
    let inStr = false;
    let esc = false;
    for (let i = start; i < t.length; i++) {
      const c = t[i];
      if (inStr) {
        if (esc) esc = false;
        else if (c === "\\") esc = true;
        else if (c === '"') inStr = false;
        continue;
      }
      if (c === '"') inStr = true;
      else if (c === open) depth++;
      else if (c === close && --depth === 0) return JSON.parse(t.slice(start, i + 1));
    }
  }
  throw new Error("Svaret indeholdt ikke gyldig JSON");
}
