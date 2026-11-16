import { MODEL_GROUPS, modelGroupOf } from "@miai/shared";
import type { CaseResult } from "./types.js";

/** Samler resultater pr. model × konfiguration og pr. kategori og skriver en Markdown-rapport (#2). */

/**
 * Kategorierne er de opgavetyper, brugeren vælger model til i Indstillinger (MODEL_GROUPS), så
 * "bedste model pr. kategori" kan sættes direkte ind.
 */
export const categoryOf = (task: string) => MODEL_GROUPS.find((g) => g.id === modelGroupOf(task))!.label.da;

export interface Aggregate {
  n: number;
  passRate: number;
  /** Andel strukturerede svar, der var gyldige i første forsøg / efter genforsøg. */
  firstTry?: number;
  afterRetry?: number;
  repetitionRate: number;
  /** Gennemsnitlig dommerscore (1–5), hvis der blev dømt. */
  judge?: number;
  medianMs: number;
  /** Tokens pr. sekund (kun hvor serveren oplyser usage). */
  tokensPerSec?: number;
  errors: number;
}

const avg = (xs: number[]) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : undefined);
const median = (xs: number[]) => {
  if (!xs.length) return 0;
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
};

export function aggregate(rs: CaseResult[]): Aggregate {
  const structured = rs.filter((r) => r.structured);
  const judged = rs.filter((r) => r.judge && r.judge.score > 0).map((r) => r.judge!.score);
  const tps = rs.filter((r) => r.completionTokens && r.durationMs).map((r) => r.completionTokens! / (r.durationMs / 1000));
  return {
    n: rs.length,
    passRate: rs.filter((r) => r.ok).length / Math.max(1, rs.length),
    firstTry: structured.length ? structured.filter((r) => r.firstTryValid).length / structured.length : undefined,
    afterRetry: structured.length ? structured.filter((r) => r.validAfterRetry).length / structured.length : undefined,
    repetitionRate: rs.filter((r) => r.repetition).length / Math.max(1, rs.length),
    judge: avg(judged),
    medianMs: median(rs.map((r) => r.durationMs)),
    tokensPerSec: avg(tps),
    errors: rs.filter((r) => r.error).length,
  };
}

function groupBy<T>(xs: T[], key: (x: T) => string): Map<string, T[]> {
  const m = new Map<string, T[]>();
  for (const x of xs) m.set(key(x), [...(m.get(key(x)) ?? []), x]);
  return m;
}

const pct = (x?: number) => (x === undefined ? "–" : `${Math.round(x * 100)} %`);
const num = (x?: number, d = 1) => (x === undefined ? "–" : x.toFixed(d).replace(".", ","));
const secs = (ms: number) => `${(ms / 1000).toFixed(1).replace(".", ",")} s`;

/** Bedste model pr. kategori: højeste beståelsesrate, derefter dommerscore, derefter hastighed. */
export function bestPerCategory(rs: CaseResult[]): { category: string; model: string; config: string; agg: Aggregate }[] {
  const out: { category: string; model: string; config: string; agg: Aggregate }[] = [];
  for (const [category, inCat] of groupBy(rs, (r) => categoryOf(r.task))) {
    const ranked = [...groupBy(inCat, (r) => `${r.model}\u0000${r.config}`)]
      .map(([k, v]) => ({ model: k.split("\u0000")[0], config: k.split("\u0000")[1], agg: aggregate(v) }))
      .sort((a, b) => b.agg.passRate - a.agg.passRate || (b.agg.judge ?? 0) - (a.agg.judge ?? 0) || a.agg.medianMs - b.agg.medianMs);
    if (ranked[0]) out.push({ category, ...ranked[0] });
  }
  return out.sort((a, b) => a.category.localeCompare(b.category, "da"));
}

export function markdownReport(rs: CaseResult[], meta: { endpoint: string; date: string }): string {
  const lines: string[] = [];
  lines.push(`# Eval-rapport ${meta.date}`, "", `Endpoint: \`${meta.endpoint}\` · ${rs.length} kørsler`, "");

  lines.push("## Overblik pr. model", "");
  lines.push("| Model | Konfiguration | Bestået | JSON 1. forsøg | JSON efter genforsøg | Gentagelser | Dommer | Median | Tokens/s | Fejl |");
  lines.push("|---|---|---|---|---|---|---|---|---|---|");
  for (const [k, v] of groupBy(rs, (r) => `${r.model}\u0000${r.config}`)) {
    const [model, config] = k.split("\u0000");
    const a = aggregate(v);
    lines.push(
      `| ${model} | ${config} | ${pct(a.passRate)} | ${pct(a.firstTry)} | ${pct(a.afterRetry)} | ${pct(a.repetitionRate)} | ${num(a.judge)} | ${secs(a.medianMs)} | ${num(a.tokensPerSec, 0)} | ${a.errors} |`,
    );
  }

  lines.push("", "## Bedste model pr. kategori", "");
  lines.push("| Kategori | Model | Konfiguration | Bestået | Dommer | Median |", "|---|---|---|---|---|---|");
  for (const b of bestPerCategory(rs))
    lines.push(`| ${b.category} | ${b.model} | ${b.config} | ${pct(b.agg.passRate)} | ${num(b.agg.judge)} | ${secs(b.agg.medianMs)} |`);

  lines.push("", "## Pr. kategori og sprog", "");
  const models = [...new Set(rs.map((r) => `${r.model} (${r.config})`))];
  lines.push(`| Kategori | Sprog | ${models.join(" | ")} |`, `|---|---|${models.map(() => "---").join("|")}|`);
  for (const [k, v] of [...groupBy(rs, (r) => `${categoryOf(r.task)}\u0000${r.lang}`)].sort(([a], [b]) => a.localeCompare(b, "da"))) {
    const [cat, lang] = k.split("\u0000");
    const cells = models.map((m) => pct(aggregate(v.filter((r) => `${r.model} (${r.config})` === m)).passRate));
    lines.push(`| ${cat} | ${lang} | ${cells.join(" | ")} |`);
  }

  const failed = rs.filter((r) => !r.ok);
  if (failed.length) {
    lines.push("", "## Fejlede cases", "");
    lines.push("| Case | Model | Konfiguration | Årsag |", "|---|---|---|---|");
    for (const r of failed) {
      const why =
        r.error ??
        r.checks
          .filter((c) => !c.ok)
          .map((c) => `${c.check}: ${c.detail ?? ""}`)
          .join("; ");
      lines.push(`| ${r.caseId} | ${r.model} | ${r.config} | ${why.replace(/\|/g, "\\|").replace(/\n/g, " ").slice(0, 200)} |`);
    }
  }
  return lines.join("\n") + "\n";
}
