import { mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import { markdownReport } from "./report.js";
import { runMatrix, type EndpointConfig } from "./runner.js";
import { EvalCase } from "./types.js";

/**
 * Kør eval-sættet mod en OpenAI-kompatibel endpoint (#2).
 *
 *   npm run eval -w @miai/evals -- --base-url http://127.0.0.1:11434/v1 --model gemma3:12b --model qwen3:14b
 *
 * Se docs/EVAL.md for alle tilvalg.
 */

const here = dirname(fileURLToPath(import.meta.url));
export const CASES_DIR = resolve(here, "../cases");

export function loadCases(dir = CASES_DIR): EvalCase[] {
  const cases = readdirSync(dir)
    .filter((f) => f.endsWith(".json"))
    .sort()
    .flatMap((f) => {
      const raw = JSON.parse(readFileSync(join(dir, f), "utf8")) as unknown[];
      return raw.map((c, i) => {
        const parsed = EvalCase.safeParse(c);
        if (!parsed.success) throw new Error(`${f}[${i}]: ${parsed.error.message}`);
        return parsed.data;
      });
    });
  const dupes = cases.map((c) => c.id).filter((id, i, all) => all.indexOf(id) !== i);
  if (dupes.length) throw new Error(`Dublerede case-id'er: ${dupes.join(", ")}`);
  return cases;
}

const HELP = `Brug: npm run eval -w @miai/evals -- --base-url <url> --model <navn> [--model <navn> …]

  --base-url <url>        OpenAI-kompatibel endpoint inkl. /v1 (fx http://127.0.0.1:11434/v1)
  --api-key <nøgle>       API-nøgle (eller MIAI_EVAL_API_KEY)
  --model <navn>          Model at evaluere; gentag for flere
  --structured <tilstand> json_schema | json_object | none; gentag for at sammenligne (standard: json_schema)
  --presence <værdi>      presence_penalty, fx 0, 1 eller 1.5; gentag for at sammenligne (standard: 0)
  --filter <tekst>        Kun cases, hvis id, opgave eller sprog indeholder teksten (fx "formula" eller "da")
  --repeat <n>            Kør hver case n gange (måler stabilitet)
  --judge-model <navn>    Dommermodel til kvalitetsscore 1–5 (bruger casens rubric)
  --judge-base-url <url>  Endpoint til dommeren (standard: --base-url)
  --out <mappe>           Hvor resultater skrives (standard: packages/evals/results)
  --list                  Vis cases og afslut`;

async function main() {
  const { values } = parseArgs({
    options: {
      "base-url": { type: "string" },
      "api-key": { type: "string" },
      model: { type: "string", multiple: true },
      structured: { type: "string", multiple: true },
      presence: { type: "string", multiple: true },
      filter: { type: "string" },
      repeat: { type: "string" },
      "judge-model": { type: "string" },
      "judge-base-url": { type: "string" },
      out: { type: "string" },
      list: { type: "boolean" },
      help: { type: "boolean", short: "h" },
    },
  });
  if (values.help) return console.log(HELP);

  let cases = loadCases();
  if (values.filter) {
    const f = values.filter.toLowerCase();
    cases = cases.filter((c) => c.id.includes(f) || c.task.includes(f) || c.lang === f);
  }
  if (values.list) {
    for (const c of cases) console.log(`${c.id.padEnd(32)} ${c.task.padEnd(18)} ${c.lang}  ${c.checks.map((x) => x.type).join(", ")}`);
    return console.log(`\n${cases.length} cases`);
  }

  const baseUrl = values["base-url"];
  if (!baseUrl || !values.model?.length) {
    console.error(HELP);
    process.exit(2);
  }
  const apiKey = values["api-key"] ?? process.env.MIAI_EVAL_API_KEY;
  const modes = (values.structured ?? ["json_schema"]) as EndpointConfig["structuredOutput"][];
  const penalties = (values.presence ?? ["0"]).map(Number);
  const configs = modes.flatMap((structuredOutput) => penalties.map((presencePenalty) => ({ baseUrl, apiKey, structuredOutput, presencePenalty })));
  const judge = values["judge-model"] ? { baseUrl: values["judge-base-url"] ?? baseUrl, apiKey, model: values["judge-model"] } : undefined;

  const results = await runMatrix(cases, {
    models: values.model,
    configs,
    judge,
    repeat: values.repeat ? Number(values.repeat) : 1,
    onResult: (r, i, total) =>
      console.log(
        `[${String(i).padStart(3)}/${total}] ${r.ok ? "✓" : "✗"} ${r.model} ${r.config} ${r.caseId} (${(r.durationMs / 1000).toFixed(1)} s)` +
          (r.ok
            ? ""
            : ` – ${
                r.error ??
                r.checks
                  .filter((c) => !c.ok)
                  .map((c) => `${c.check}: ${c.detail}`)
                  .join("; ")
              }`),
      ),
  });

  const date = new Date().toISOString().slice(0, 16).replace(/[:T]/g, "-");
  const outDir = resolve(values.out ?? resolve(here, "../results"));
  mkdirSync(outDir, { recursive: true });
  const jsonFile = join(outDir, `${date}.json`);
  const mdFile = join(outDir, `${date}.md`);
  writeFileSync(jsonFile, JSON.stringify({ endpoint: baseUrl, date, results }, null, 2));
  writeFileSync(mdFile, markdownReport(results, { endpoint: baseUrl, date }));
  console.log(`\nRapport: ${mdFile}\nRådata:  ${jsonFile}`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
