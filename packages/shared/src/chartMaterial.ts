/**
 * Chat → agent, når chatsvaret tegner en graf som Mermaid (xychart eller pie): agenten skal bruge
 * tallene, ikke diagramkoden. Koden fylder meget (tema-linjer, akser, kodehegn), og materialet til
 * agenten er begrænset, så en samtale med tre grafer blev skåret af midt i den anden – agenten
 * manglede så tal og fandt på at hente dem selv. Her erstattes hver graf med dens tal i kort form.
 *
 * Tekst uden for graferne og andre figurer (rutediagrammer osv.) står uændret.
 */

const FENCE = /```mermaid[^\n]*\n([\s\S]*?)(?:```|$)/g;

/** Indholdet af en [a, b, "c d"]-liste, med anførselstegn fjernet. */
function listItems(src: string): string[] {
  const items = src.match(/"[^"]*"|[^,\s][^,]*/g) ?? [];
  return items.map((x) => x.trim().replace(/^"|"$/g, "")).filter(Boolean);
}

function chartData(code: string): string | undefined {
  const lines = code
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l && !l.startsWith("%%"));
  const kind = lines[0]?.match(/^(xychart(?:-beta)?|pie)\b/)?.[1];
  if (!kind) return undefined;
  const title = lines.map((l) => l.match(/^(?:pie\s+(?:showData\s+)?)?title\s+"?([^"]+)"?\s*$/)?.[1]).find(Boolean);
  const out = [`Chart${title ? `: ${title.trim()}` : ""}`];
  if (kind === "pie") {
    for (const l of lines) {
      const m = l.match(/^"([^"]+)"\s*:\s*(-?[\d.]+)\s*$/);
      if (m) out.push(`${m[1]}\t${m[2]}`);
    }
  } else {
    let series = 0;
    for (const l of lines) {
      const x = l.match(/^x-axis\s+(?:"[^"]*"\s+)?\[(.*)\]\s*$/);
      const y = l.match(/^y-axis\s+"([^"]+)"/);
      const s = l.match(/^(line|bar)\s+(?:"([^"]*)"\s+)?\[(.*)\]\s*$/);
      if (x) out.push(`x (${listItems(x[1]).length}): ${listItems(x[1]).join(", ")}`);
      else if (y) out.push(`y: ${y[1]}`);
      else if (s) {
        const values = listItems(s[3]);
        out.push(`series ${++series}${s[2] ? ` "${s[2]}"` : ""} (${values.length}): ${values.join(", ")}`);
      }
    }
  }
  return out.length > 1 ? out.join("\n") : undefined;
}

/** Chatsvaret som materiale til agenten: grafer af tal er erstattet af deres tal. */
export function chartMaterial(answer: string): string {
  return answer.replace(FENCE, (whole, code: string) => {
    const data = chartData(code);
    return data ? `${data}\n` : whole;
  });
}
