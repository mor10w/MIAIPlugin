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

- **`tasks/`** (samlet eksport i `tasks.ts`): opgavekatalog. `catalog.ts` har opgaverne, hver med sin prompt-bygger, output-type (tekst eller JSON med Zod-skema), parametre og indsætningsmåder. `definition.ts` har typerne, `prompts.ts` de fælles prompt-byggesten og `results.ts` Zod-skemaerne for strukturerede svar. Dokumentindhold pakkes i `<document>`-tags og behandles som data.
- **`userTasks.ts`**: brugerdefinerede opgaver (F14, ADR-0011) – `UserTaskFile`-skema (JSON-filer med prompt-skabelon), `renderTemplate`, omsætning til `TaskDefinition` og det serialiserbare `TaskView`.
- **`agent.ts`**: agent-tilstand (#6, ADR-0012) – `AgentPlan`/`AgentAction` (lukket union af Word- og Excel-handlinger), værktøjsbeskrivelse til prompten og `checkPlan`, der markerer handlinger, som ikke kan udføres. Udførelsen sker i add-in'ets host-adaptere (`applyAgent`) med fortryd for hele sessionen.
- **`design.ts`**: præsentationsdesign (ADR-0014) – `DesignPlan` (stil, palet, skrifttyper, slides på syv layouttyper), kontrastretning (`fixPalette`) og designregler til modellen. Add-in'et bygger masteren med `deck/templates.ts` og `deck/buildDeck.ts` (pptxgenjs) og sætter den ind med `insertSlidesFromBase64`.
- **`knowledge.ts`**: lokal RAG (F15, ADR-0013) – opdeling i uddrag (`chunkText`), normalisering og brute force-søgning (`topK`). `openai.ts` har `embed()` og `embeddingPrefix()`; chat-opgaven giver modellen uddrag som nummererede `<source>`-blokke.
- **`audit.ts`**: `AuditEntry`-skema (kun metadata) og grænser for loggen.
- **`engine.ts`** (`executeTask`):
  - vælger model (forespørgsel → model pr. opgave → standard → første fra `/v1/models`)
  - tjekker politikken
  - streamer svaret og filtrerer `<think>`
  - ved JSON-opgaver: sætter `response_format` efter endpointens `structuredOutput`, sender skemaet i prompten, validerer med Zod og prøver én gang mere ved fejl
  - sætter faste instruktioner (#78) ind i systembeskeden for alle opgaver: organisationens (`Policy.instructions`) før brugerens (`Settings.instructions`) for værten i `req.context.host`. Indpakningen (`instructionsBlock` i `tasks/prompts.ts`) står efter grundreglerne og før alt dokumentindhold
- **`context.ts`** (#3): token-estimat, svarbudget, opdeling ved afsnit og detektion af afkortning. Motoren bruger dem til:
  - at begrænse `max_tokens` og afvise for lange prompts, når kontekstlængden er kendt (endpointens `contextTokens`, aflæst fra `/v1/models` eller fra værktøjets eget API, se `modelInfo.ts`)
  - at køre map-reduce for opgaver med `chunking: "map-reduce"` (opsummering)
  - at melde `done.warnings`, når `usage.prompt_tokens` viser, at modellen fik mindre, end der blev sendt, eller når `finish_reason` er `length`

  Tekst fra en værktøjsrunde trækkes tilbage med en `retract`-event.
- **`modelInfo.ts`** (#73): modellens kontekst og kapabiliteter fra værktøjets eget API, fordi OpenAI's `/v1/models` hverken oplyser Ollamas kontekst eller LM Studios indlæste kontekst. Ollama: `POST /api/show` (Modelfile-parametre, trænet kontekst, `capabilities`) og `GET /api/ps` (konteksten, en indlæst model kører med). LM Studio: `GET /api/v1/models` (`max_context_length`, den indlæste instans' `context_length`, `vision`, `trained_for_tool_use`). Kaldene går til endpointens egen host og port. En server, der ikke svarer som nogen af dem, spørges ikke igen i 10 minutter. `engine.ts` samler det i `modelDetails()` (30 s cache) og melder `small-context` i `done.warnings`, når agent eller chat kører med under 16.384 tokens.
- **`openai.ts`**: minimal klient. Bruger `fetch` + `ReadableStream.getReader()` og har fallback for `AbortSignal.any`. GET-kald sendes uden ekstra headers, så browseren undgår CORS-preflight. Tåler servere, der ignorerer `stream: true`.
- **`policy.ts`**: `allowedHosts` med `*.domæne`-wildcards.

### 2. Office Add-in (`packages/addin`)

Host-adaptere: Word, Excel, PowerPoint og Outlook (`hosts/outlook.ts`, Mailbox-API'et i læse- og skrivetilstand, eget MailApp-manifest – ADR-0015).
- Task pane (React 19 + Fluent UI v9) med fanerne Handlinger, Opgaver, Chat og Indstillinger.
- **Backends** (`src/backend`) med samme interface:
  - `DirectBackend`: kører `executeTask` i webviewet. Indstillinger og politik gemmes i `localStorage` med fallback til hukommelse; brugerdefinerede opgaver (F14) og audit-loggen (ringbuffer, maks. 200 poster) ligeså. Ved forbindelsesfejl vises præcis CORS-vejledning for add-in'ets origin; har browseren afvist lokal netværksadgang (Chrome og Edge, `localNetwork.ts`), vises det i stedet. Ved central udrulning henter den `miai-config.json` fra add-in'ets egen adresse: organisationens låste politik og standardindstillinger (ADR-0015).
  - `BridgeBackend`: kalder `/api/*` på samme origin og læser SSE-strømmen.
- **Regnearksfunktioner** (`src/cell`, #80, ADR-0018): `=MIAI.ASK()` og `=MIAI.MODELS()` i Excel. `functions.ts` registrerer dem med `CustomFunctions.associate` og kalder opgaven `cell-ask` gennem backend'en; `queue.ts` er køen med loft og Stop, `cache.ts` den vedvarende cache, og `value.ts` gør svaret til en celleværdi. De kræver Excels shared runtime, som manifesterne til Word, Excel og PowerPoint slår til i Excel (bridge, dev og hostet).
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
- **Excel**: formlen skrives i cellen, og resultatet læses, så fejl som #NAME? meldes. Tidligere formler gemmes, så **Fortryd** virker (Ctrl+Z fortryder ikke altid add-in-ændringer). Oversættelse af områder sker som TSV. Formelsporing (#76): add-in'et gennemløber Excels afhængighedsgraf (`getDirectPrecedents`/`getDirectDependents`) med loft over dybde og antal celler og giver modellen en kompakt liste; den rene logik ligger i `hosts/excelTrace.ts`, Office.js-delen i `hosts/excelTraceReader.ts` (ADR-0012).
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
- **vidensindeks** (F15, ADR-0013): `knowledge/` udtrækker tekst (docx/pdf/xlsx/pptx/md), indekserer brugerens mapper med embeddings i `knowledge.db` (node:sqlite), holder indekset opdateret (fs.watch) og søger med brute force. `GET /api/knowledge`, `POST /api/knowledge/reindex`, `GET /api/knowledge/chunks/:id`, `POST /api/knowledge/open`. Mapper tjekkes mod `policy.knowledgeFolders`.
- **audit-log** med metadata (aldrig indhold): JSONL i `audit.log` med rotation ved 1 MB (`audit.log.old`)
- UI uden internet
- filsystemadgang til senere RAG (fase 4)

Sikkerhed:
- lytter kun på loopback
- `Host`-tjek mod DNS rebinding
- kræver `X-MIAI`-header og eget `Origin` (holder fremmede websider ude)
- kræver adgangstoken pr. bruger i `X-MIAI-Token` (#11, holder andre lokale processer og brugere ude). Tokenet ligger i `bridge.json` (0600) og i brugerens personlige manifest (`…/taskpane.html?token=…`), ikke i den serverede side. Porten vælges pr. bruger og huskes.
- API-nøgler ligger i OS'ets nøglering (#16, `@napi-rs/keyring`, tjeneste `MIAI`, konto `endpoint:<id>`); `config.json` har kun referencen `keyring:<id>`. Uden nøglering (eller med `MIAI_KEYRING=off`) bliver nøglerne i `config.json` (0600) med en advarsel. `GET /api/settings` maskerer altid nøglerne.

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
| hosted | GitHub Pages (`.github/workflows/pages.yml`) | Direkte | `<pages-url>/manifest.xml` og `<pages-url>/manifest-outlook.xml` |
| dev | Vite på `https://localhost:3000` | Direkte | `packages/addin/manifest.dev.xml` og `manifest.outlook.dev.xml` |
| bridge | MIAI Bridge på `https://localhost:44321` | Bridge | `packages/addin/manifest.xml` og `manifest.outlook.xml` (personlige udgaver med token i konfigurationsmappen) |

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
│  ├─ bridge/       # valgfri lokal service
│  ├─ app/          # MIAI App (Electron, Mac og Windows; platform.js, opdatering i updates.js/updateState.js, udviklerudgave i devBuild.js – ADR-0016)
│  └─ evals/        # eval-sæt og harness mod en OpenAI-endpoint (#2)
├─ scripts/         # manifester, ikoner, sideload (Mac/Windows)
├─ .github/workflows/  # CI og GitHub Pages
└─ docs/            # plan, arkitektur, udvikling, ADR'er
```
