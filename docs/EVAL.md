# Eval-sæt og modelvalg (#2)

`packages/evals` evaluerer MIAI's indbyggede opgaver mod en hvilken som helst OpenAI-kompatibel endpoint (Ollama, LM Studio, llama.cpp, vLLM, oMLX, en intern gateway …).

Harnesset kalder `executeTask` fra `@miai/shared`. Det er samme motor, som add-in'et og bridgen bruger, så det er produktets egne prompts, validering, genforsøg, `<think>`-filter og sløjfe-vagt, der evalueres. Der er ingen separat kopi af prompts at holde i sync.

## Indhold

57 cases i `packages/evals/cases/*.json`, på dansk og engelsk og fordelt på alle 18 indbyggede opgaver (F1–F13, agent-tilstand, præsentationsdesign, svar på mail og web-søgning):

| Fil | Opgaver | Hvad der især testes |
|---|---|---|
| `skrivning.json` | omskriv, parafrasér, forenkl, artikel, salgstekst, disposition | tone, sprog, fakta bevaret, intet forord ("Her er …"), længde i forhold til input |
| `oversaettelse.json` | oversæt (da↔en, da→de, Excel-celler) | målsprog, tal og datoer i dansk format, TSV-form bevaret |
| `resume.json` | opsummér, analysér data | punkter/længde, nøgletal med, **prompt injection** i dokumentet |
| `struktureret.json` | grammatik, slides, diagram | gyldigt JSON i første forsøg, antal slides/noder, ingen opdigtede rettelser i korrekt tekst |
| `formler.json` | lav og forklar formler | `SUMIFS`/`XLOOKUP`/`COUNTIFS`, engelske funktionsnavne, komma som separator |
| `chat.json` | chat og web-søgning | svar ud fra dokumentet, "står ikke i dokumentet" (hallucination), Excel-omfang, samtalehistorik, kalder `web_search` når relevant – og lader være, når dokumentet svarer; henviser til agenten (ingen VBA) ved ønsker om formatering |
| `agent.json` | agent i Word, Excel og PowerPoint | gyldige planer (afsnit/figurer/tabeller findes), formatering bevares: `setStyle` frem for slet+indsæt, korte erstatninger, `appendRows` i Excel-tabeller, tomme rækker og tabeltypografi i Excel, prompt injection |

Hver case har automatiske tjek og en `rubric` til en valgfri dommermodel. Formatet er beskrevet i `packages/evals/src/types.ts`.

## Kørsel

Kræver en kørende endpoint med modellerne hentet. Eksempel med Ollama:

```bash
npm run eval -- --base-url http://127.0.0.1:11434/v1 --model gemma3:12b --model qwen3:14b --model mistral-small3.2
```

Laptop-kandidater, kun formelopgaverne, tre gentagelser for at måle stabilitet:

```bash
npm run eval -- --base-url http://127.0.0.1:11434/v1 --model qwen3:4b --model gemma3:4b --model phi4-mini --filter formula --repeat 3
```

Struktureret output og `presence_penalty` kan sammenlignes i samme kørsel. Hver kombination bliver en konfiguration i rapporten:

```bash
npm run eval -- --base-url http://127.0.0.1:1234/v1 --model qwen3-14b --structured json_schema --structured json_object --structured none --presence 0 --presence 1 --presence 1.5
```

Med en dommermodel får hvert svar også en kvalitetsscore fra 1 til 5 efter casens kriterier. Brug en større model end dem, der evalueres:

```bash
npm run eval -- --base-url http://127.0.0.1:11434/v1 --model qwen3:4b --judge-model qwen3:32b
```

Øvrige tilvalg: `--api-key` (eller `MIAI_EVAL_API_KEY`), `--judge-base-url`, `--out` og `--list`. Kør `npm run eval -- --help` for at se dem alle.

Resultatet skrives til `packages/evals/results/<tidspunkt>.md` (rapport) og `.json` (rådata med hvert svar). Mappen er ignoreret af git. Commit de rapporter, der danner grundlag for en anbefaling, under `docs/eval/`.

## Hvad rapporten måler

| Måling | Betydning |
|---|---|
| Bestået | Andel cases, hvor alle automatiske tjek holdt |
| JSON 1. forsøg / efter genforsøg | Strukturerede opgaver: gyldigt svar uden og med motorens ene genforsøg. Måles pr. `structuredOutput`-tilstand |
| Gentagelser | Andel svar, sløjfe-vagten afbrød (`stopReason: "repetition"`). Sammenlign `--presence 0/1/1.5`, og tjek sproget på dansk ved 1,5 |
| Dommer | Gennemsnitlig score (1–5), hvis `--judge-model` er sat |
| Median, tokens/s | Svartid og hastighed (tokens/s kun, hvor serveren oplyser usage) |

Rapporten har også **bedste model pr. kategori** (Skrivning, Oversættelse, Resumé og analyse, Struktureret, Formler, Chat). Det er grundlaget for anbefalinger pr. opgave i UI'et (#13).

Web-søgning bruger faste søgeresultater. Der sendes intet på nettet; det er kun modellens beslutning om at søge, der måles.

## Kandidater

Fra PLAN.md §5 og tilføjelserne i #2:

| Brug | Modeller |
|---|---|
| Standard, dansk | Gemma 3 12B, Qwen3 14B, Mistral Small 3.x |
| Laptop uden GPU | Qwen3 4B, Gemma 3 4B, Phi-4-mini |
| Formler, struktureret output | Qwen3, Qwen2.5-Coder |
| Mac (MLX/oMLX) | MLX-udgaver af ovenstående |
| GPU-server (organisationer) | Gemma 3 27B, Qwen3 32B |

## Anbefalinger

*Udfyldes efter første kørsel på rigtig hardware; fremgangsmåden står i [TESTVEJLEDNING.md](TESTVEJLEDNING.md#2-eval-og-modelvalg).* Indtil da gælder forslagene i PLAN.md §5. Notér hardware, værktøj og version, kvantisering og dato sammen med rapporten.

## Udvid sættet

Tilføj en case i den relevante fil i `cases/` med et unikt `id`, en `task` fra opgavekataloget, `lang`, `context` (som add-in'et ville sende det), `params`, `checks` og `rubric`. `npm test -w @miai/evals` validerer, at alle cases kan parses, at id'erne er unikke, og at alle indbyggede opgaver er dækket.
