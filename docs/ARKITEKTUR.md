# Arkitektur

## Overblik

```
                  ┌────────────── packages/shared (@miai/shared) ──────────────┐
                  │ types · tasks (prompts, Zod-skemaer) · engine (executeTask) │
                  │ openai (klient) · think (filter) · json · policy · text     │
                  └──────────────┬───────────────────────────┬─────────────────┘
                                 │ bruges af                 │ bruges af
          ┌──────────────────────▼─────────┐      ┌──────────▼──────────────────┐
          │ packages/addin (task pane)     │      │ packages/bridge (valgfri)   │
          │ host-adaptere · UI · backend:  │      │ Fastify, HTTPS på loopback  │
          │  ├─ DirectBackend ─────────────┼──┐   │ serverer add-in + /api/*    │
          │  └─ BridgeBackend ─────────────┼──┼──▶│ config.json, maskinpolitik, │
          └────────────────────────────────┘  │   │ audit-log                   │
                                              │   └──────────┬──────────────────┘
                                              ▼              ▼
                                OpenAI-kompatibel endpoint (/v1/models, /v1/chat/completions)
                                Ollama · LM Studio · llama.cpp · vLLM · LocalAI · Jan · …
```

Tilstanden vælges ved runtime (`backend/index.ts`). Hvis task pane serveres af en MIAI Bridge (`/api/health` svarer på samme origin), bruges den. Ellers bruges direkte forbindelse. Se [ADR-0007](adr/0007-bridge-valgfri-direkte-forbindelse.md).

## Komponenter

### 1. Fælles motor (`packages/shared`)
Ren TypeScript uden Node-afhængigheder, så den kører i Office-webviewet og i Node.

- **`tasks.ts`**: opgavekatalog. Hver opgave har prompt-bygger, output-type (tekst eller JSON med Zod-skema), parametre og indsætningsmåder. Dokumentindhold pakkes i `<document>`-tags og behandles som data.
- **`userTasks.ts`**: brugerdefinerede opgaver (F14, ADR-0011) – `UserTaskFile`-skema (JSON-filer med prompt-skabelon), `renderTemplate`, omsætning til `TaskDefinition` og det serialiserbare `TaskView`.
- **`audit.ts`**: `AuditEntry`-skema (kun metadata) og grænser for loggen.
- **`engine.ts`** (`executeTask`):
  - vælger model (forespørgsel → model pr. opgave → standard → første fra `/v1/models`)
  - tjekker politikken
  - streamer svaret og filtrerer `<think>`
  - ved JSON-opgaver: sætter `response_format` efter endpointens `structuredOutput`, sender skemaet i prompten, validerer med Zod og prøver én gang mere ved fejl
- **`context.ts`** (#3): token-estimat, svarbudget, opdeling ved afsnit og detektion af afkortning. Motoren bruger dem til:
  - at begrænse `max_tokens` og afvise for lange prompts, når kontekstlængden er kendt (endpointens `contextTokens` eller aflæst fra `/v1/models`)
  - at køre map-reduce for opgaver med `chunking: "map-reduce"` (opsummering)
  - at melde `done.warnings`, når `usage.prompt_tokens` viser, at modellen fik mindre, end der blev sendt, eller når `finish_reason` er `length`

  Tekst fra en værktøjsrunde trækkes tilbage med en `retract`-event.
- **`openai.ts`**: minimal klient. Bruger `fetch` + `ReadableStream.getReader()` og har fallback for `AbortSignal.any`. GET-kald sendes uden ekstra headers, så browseren undgår CORS-preflight. Tåler servere, der ignorerer `stream: true`.
- **`policy.ts`**: `allowedHosts` med `*.domæne`-wildcards.

### 2. Office Add-in (`packages/addin`)
- Task pane (React 19 + Fluent UI v9) med fanerne Handlinger, Opgaver, Chat og Indstillinger.
- **Backends** (`src/backend`) med samme interface:
  - `DirectBackend`: kører `executeTask` i webviewet. Indstillinger og politik gemmes i `localStorage` med fallback til hukommelse; brugerdefinerede opgaver (F14) og audit-loggen (ringbuffer, maks. 200 poster) ligeså. Ved forbindelsesfejl vises præcis CORS-vejledning for add-in'ets origin.
  - `BridgeBackend`: kalder `/api/*` på samme origin og læser SSE-strømmen.
- **Host-adaptere**:

```ts
interface HostAdapter {
  host: "word" | "excel" | "powerpoint";
  supports(mode: ApplyMode): boolean;
  getContext(opts: { wholeDocumentFallback: boolean }): Promise<DocContext>;
  applyText(text, format, mode, ctx): Promise<ApplyOutcome>;   // replace | trackChanges | insertAfter | comment | cells
  applyGrammar?(result, mode, ctx): Promise<ApplyOutcome>;     // Word: search + sporede ændringer
  applySlides?(result): Promise<ApplyOutcome>;                 // PowerPoint
  applyFormula?(result, ctx): Promise<ApplyOutcome>;           // Excel, med validering og fortryd
}
```

- **Word**: markeringen gemmes som sporet objekt (`trackedObjects`), så resultatet lander det rigtige sted, selvom brugeren klikker andetsteds imens. Grammatik bruger `search` og sporede ændringer.
- **Excel**: formlen skrives i cellen, og resultatet læses, så fejl som #NAME? meldes. Tidligere formler gemmes, så **Fortryd** virker (Ctrl+Z fortryder ikke altid add-in-ændringer). Oversættelse af områder sker som TSV.
- **PowerPoint**: tekst via `get/setSelectedDataAsync`. Slides oprettes med `slides.add`, og layoutets pladsholdere udfyldes; findes de ikke, tilføjes tekstbokse (PowerPointApi 1.4). Plan B ved firmaskabeloner: `.pptx` + `insertSlidesFromBase64`.
- **Sikkerhed**: modeloutput renses med DOMPurify før visning og indsættelse som HTML.

### 3. MIAI Bridge (`packages/bridge`) – valgfri
Node.js-service (Fastify) på `https://127.0.0.1:44321`. Serverer add-in'ets filer og API'et fra samme origin.

| Endpoint | Formål |
|----------|--------|
| `GET  /api/health` | Status og om den aktive endpoint svarer |
| `GET  /api/models` | Modeller fra endpointens `/v1/models` |
| `GET  /api/tasks` | Opgavekatalog (indbygget + brugerdefineret) |
| `POST /api/tasks/{id}` | Kør opgave (indbygget eller brugerdefineret); svar som SSE (`TaskEvent`) |
| `GET  /api/user-tasks` | Bruger- og maskin-opgaver med oprindelse |
| `PUT/DELETE /api/user-tasks/{id}` | Opret/ændre/slet bruger-opgave (maskin-opgaver afvises med 409) |
| `GET  /api/audit` | Seneste audit-logposter (metadata) |
| `GET/PUT /api/settings` | Indstillinger og politik (API-nøgler maskeres) |

Bridgen bidrager med det, en browser ikke kan:
- en **maskinpolitik**, som brugeren ikke kan ændre (ADR-0006)
- **brugerdefinerede opgaver** som filer: brugers mappe (`settings.customTasksDir`, standard `<config>/tasks`) og organisationsmappe (`MIAI_TASKS_DIR`), hvor maskin-opgaver vinder ved id-kollision og er read-only (ADR-0011)
- **audit-log** med metadata (aldrig indhold): JSONL i `audit.log` med rotation ved 1 MB (`audit.log.old`)
- UI uden internet
- filsystemadgang til senere RAG (fase 4)

Sikkerhed:
- lytter kun på loopback
- `Host`-tjek mod DNS rebinding
- kræver `X-MIAI`-header og eget `Origin` (holder fremmede websider ude)
- kræver adgangstoken pr. bruger i `X-MIAI-Token` (#11, holder andre lokale processer og brugere ude). Tokenet ligger i `bridge.json` (0600) og i brugerens personlige manifest (`…/taskpane.html?token=…`), ikke i den serverede side. Porten vælges pr. bruger og huskes.

### 4. Inference-værktøj (eksternt)
Et hvilket som helst OpenAI-kompatibelt værktøj. Ved direkte forbindelse skal værktøjet tillade CORS fra add-in'ets origin:

| Værktøj | CORS |
|---|---|
| Ollama | `https://localhost:*` tilladt som standard; andre origins: `OLLAMA_ORIGINS=https://mor10w.github.io` |
| LM Studio | Developer/Server → "Enable CORS" |
| vLLM | `--allowed-origins '["https://…"]'` (standard `*`) |
| llama.cpp server | Tillader som udgangspunkt alle origins |
| Andre | Se værktøjets dokumentation, eller brug bridgen |

## Hosting og manifester
Manifester genereres fra `packages/addin/manifest.template.xml` af `scripts/make-manifest.mjs`. Hver variant har eget Id:

| Variant | Add-in hentes fra | Tilstand | Fil |
|---|---|---|---|
| hosted | GitHub Pages (`.github/workflows/pages.yml`) | Direkte | `<pages-url>/manifest.xml` |
| dev | Vite på `https://localhost:3000` | Direkte | `packages/addin/manifest.dev.xml` |
| bridge | MIAI Bridge på `https://localhost:44321` | Bridge | `packages/addin/manifest.xml` |

En organisation kan hoste `dist/` på en intern HTTPS-server og generere sit eget manifest med `node scripts/make-manifest.mjs hosted https://intern.server/miai/ manifest.xml`.

## Dataflow – eksempel "Forklar formel" (Excel, direkte)
1. Brugeren markerer `C5` og klikker *Forklar formel*.
2. `ExcelAdapter.getContext()` henter formel, adresse og overskrifter/første rækker.
3. `DirectBackend.runTask` → `executeTask` bygger prompten og streamer fra `http://127.0.0.1:11434/v1/chat/completions`.
4. Task pane viser forklaringen; valgfrit *Som kommentar* på cellen.

## Repo-struktur
```
MIAIPlugin/
├─ packages/
│  ├─ shared/       # fælles motor: typer, opgaver/prompts, OpenAI-klient, politik
│  ├─ addin/        # Office.js task pane, backends, host-adaptere, manifest-skabelon
│  └─ bridge/       # valgfri lokal service
├─ scripts/         # manifester, ikoner, sideload (Mac/Windows)
├─ .github/workflows/  # CI og GitHub Pages
└─ docs/            # plan, arkitektur, udvikling, ADR'er
```
