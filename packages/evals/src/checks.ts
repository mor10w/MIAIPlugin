import { AgentPlan, checkPlan, type DocContext, type HostKind } from "@miai/shared";
import type { Check, CheckResult } from "./types.js";

/** Automatiske tjek af et svar (#2). Rene funktioner, så de kan unit-testes uden model. */

export interface Answer {
  /** Svarets tekst (for strukturerede opgaver: den rå JSON). */
  output: string;
  /** Det validerede strukturerede svar, hvis opgaven har et. */
  result?: unknown;
  /** Inputtets tekst – til længde i forhold til input. */
  input: string;
  toolCalls: string[];
  /** Værten – til agent-planer. */
  host?: HostKind;
  /** Word: dokumentets afsnit, som modellen så dem – til tjek af afsnitsnumre i agent-planer. */
  paragraphs?: DocContext["paragraphs"];
  /** Excel: arkets tabeller – til tjek af tabelnavne i agent-planer. */
  excel?: DocContext["excel"];
  /** Kodetilstanden er slået til i casen. */
  scripts?: boolean;
}

const STOPWORDS: Record<"da" | "en" | "de", string[]> = {
  da: ["og", "er", "det", "at", "en", "til", "af", "på", "med", "ikke", "som", "der", "har", "for", "vi", "jeg", "kan", "være", "også", "eller"],
  en: ["the", "and", "is", "of", "to", "in", "that", "it", "with", "not", "for", "are", "this", "be", "on", "we", "can", "also", "or", "have"],
  de: ["und", "ist", "die", "der", "das", "nicht", "mit", "zu", "auf", "ein", "eine", "für", "sie", "wir", "auch", "oder", "sind", "von", "werden", "es"],
};

export const words = (s: string) => s.match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) ?? [];

/** Sproget med flest stopord; æ/ø/å tæller med for dansk. */
export function detectLanguage(text: string): "da" | "en" | "de" | "unknown" {
  const w = words(text.toLowerCase());
  const score = (lang: "da" | "en" | "de") => w.filter((x) => STOPWORDS[lang].includes(x)).length;
  const scores = { da: score("da") + (text.match(/[æøå]/gi)?.length ?? 0) / 4, en: score("en"), de: score("de") + (text.match(/[äöüß]/gi)?.length ?? 0) / 4 };
  const best = (Object.entries(scores) as ["da" | "en" | "de", number][]).sort((a, b) => b[1] - a[1]);
  return best[0][1] === 0 || best[0][1] === best[1][1] ? "unknown" : best[0][0];
}

/** Slår en sti som "slides", "slides.0.title" eller "formula" op. */
export function getPath(value: unknown, path: string): unknown {
  return path.split(".").reduce<unknown>((v, key) => (v && typeof v === "object" ? (v as Record<string, unknown>)[key] : undefined), value);
}

const PREAMBLE = /^\s*(her er|her kommer|selvfølgelig|sure|here is|here's|certainly|of course)\b/i;

export function runCheck(check: Check, a: Answer): CheckResult {
  const text = a.output;
  const lower = text.toLowerCase();
  const r = (ok: boolean, detail?: string): CheckResult => ({ check: check.type, ok, ...(ok || !detail ? {} : { detail }) });
  switch (check.type) {
    case "contains": {
      const missing = check.values.filter((v) => !lower.includes(v.toLowerCase()));
      return check.any
        ? r(missing.length < check.values.length, `ingen af ${check.values.join(", ")}`)
        : r(missing.length === 0, `mangler ${missing.join(", ")}`);
    }
    case "notContains": {
      const found = check.values.filter((v) => lower.includes(v.toLowerCase()));
      return r(found.length === 0, `indeholder ${found.join(", ")}`);
    }
    case "regex": {
      const hit = new RegExp(check.pattern, check.flags ?? "i").test(text);
      return r(check.negate ? !hit : hit, `${check.negate ? "matcher" : "matcher ikke"} /${check.pattern}/`);
    }
    case "language": {
      const lang = detectLanguage(text);
      return r(lang === check.lang, `sprog: ${lang}`);
    }
    case "length": {
      const n = words(text).length;
      const ratio = n / Math.max(1, words(a.input).length);
      const problems = [
        check.minWords !== undefined && n < check.minWords && `${n} < ${check.minWords} ord`,
        check.maxWords !== undefined && n > check.maxWords && `${n} > ${check.maxWords} ord`,
        check.maxRatio !== undefined && ratio > check.maxRatio && `forhold ${ratio.toFixed(2)} > ${check.maxRatio}`,
        check.minRatio !== undefined && ratio < check.minRatio && `forhold ${ratio.toFixed(2)} < ${check.minRatio}`,
      ].filter(Boolean);
      return r(problems.length === 0, problems.join("; "));
    }
    case "bullets": {
      const n = text.split("\n").filter((l) => /^\s*([-*•]|\d+[.)])\s+/.test(l)).length;
      return r((check.min === undefined || n >= check.min) && (check.max === undefined || n <= check.max), `${n} punkter`);
    }
    case "noPreamble": {
      const fenced = /^\s*```/.test(text);
      return r(!PREAMBLE.test(text) && !fenced, fenced ? "kodeblok rundt om svaret" : `starter med "${text.trim().slice(0, 30)}"`);
    }
    case "result": {
      const v = getPath(a.result, check.path);
      if (v === undefined) return r(false, `${check.path} mangler`);
      if (Array.isArray(v)) {
        const ok = (check.minItems === undefined || v.length >= check.minItems) && (check.maxItems === undefined || v.length <= check.maxItems);
        return r(ok, `${check.path}: ${v.length} elementer`);
      }
      if (check.pattern) return r(new RegExp(check.pattern, check.flags ?? "").test(String(v)), `${check.path} = ${String(v).slice(0, 80)}`);
      return r(true);
    }
    case "edits": {
      const edits = ((a.result as { edits?: { original: string; replacement: string }[] })?.edits ?? []).map((e) => e.replacement.toLowerCase());
      const missing = check.expect.filter((e) => !edits.some((rep) => rep.includes(e.fix.toLowerCase())));
      const tooMany = check.maxEdits !== undefined && edits.length > check.maxEdits;
      return r(
        missing.length === 0 && !tooMany,
        [missing.length && `mangler ${missing.map((m) => m.fix).join(", ")}`, tooMany && `${edits.length} rettelser`].filter(Boolean).join("; "),
      );
    }
    case "tsv": {
      const rows = text.trim().split("\n");
      const colsOk = check.columns === undefined || rows.every((row) => row.split("\t").length === check.columns);
      const rowsOk = check.rows === undefined || rows.length === check.rows;
      return r(colsOk && rowsOk, `${rows.length} rækker, kolonner ${[...new Set(rows.map((row) => row.split("\t").length))].join("/")}`);
    }
    case "plan": {
      const parsed = AgentPlan.safeParse(a.result);
      if (!parsed.success) return r(false, "ingen gyldig plan");
      const kinds = parsed.data.actions.map((x) => x.action as string);
      const problems = [
        check.min !== undefined && kinds.length < check.min && `${kinds.length} < ${check.min} handlinger`,
        check.max !== undefined && kinds.length > check.max && `${kinds.length} > ${check.max} handlinger`,
        ...(check.include ?? []).filter((k) => !kinds.includes(k)).map((k) => `mangler ${k}`),
        ...(check.exclude ?? []).filter((k) => kinds.includes(k)).map((k) => `indeholder ${k}`),
      ];
      if (check.valid) {
        const bad = checkPlan(parsed.data, { host: a.host ?? "word", text: a.input, paragraphs: a.paragraphs, excel: a.excel }, { scripts: a.scripts }).filter(
          (c) => c.problem,
        );
        if (bad.length) problems.push(`${bad.length} handlinger kan ikke udføres (${bad.map((b) => b.problem).join(", ")})`);
      }
      const list = problems.filter(Boolean);
      return r(list.length === 0, list.join("; "));
    }
    case "tool": {
      const called = a.toolCalls.length > 0;
      return r(called === check.called, called ? `kaldte ${a.toolCalls.join(", ")}` : "kaldte ikke web_search");
    }
  }
}
