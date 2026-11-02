# Arkitektur (udkast)

## Komponenter

### 1. Office Add-in (`packages/addin`)
- Office Web Add-in med task pane + ribbon-kommandoer (function commands).
- **Host-adaptere** skjuler forskelle mellem Word, Excel og PowerPoint bag ét interface:

```ts
interface HostAdapter {
  host: "word" | "excel" | "powerpoint" | "outlook";
  getSelection(): Promise<DocContext>;        // tekst, formel, range-adresse, slide-id …
  getDocument(opts?: { maxChars?: number }): Promise<DocContext>;
  apply(result: AiResult, mode: ApplyMode): Promise<UndoHandle>;
}
type ApplyMode = "replace" | "insertAfter" | "comment" | "trackChanges" | "newSlides" | "cells";
```

- Ændringer vises som diff i task pane før de anvendes. Word: `changeTrackingMode`
  slås midlertidigt til, så brugeren kan acceptere/afvise ændringer i Word.
- Excel: formler skrives aldrig blindt – foreslået formel valideres ved at skrive
  den til en midlertidig celle og læse resultat/fejl, før brugeren accepterer.
- PowerPoint: tekst → slides via shapes-API'et (`slides.add` + tekstbokse i
  layoutets placeholders) eller, hvis API-niveauet ikke rækker, en `.pptx`
  genereret i bridgen og indsat med `insertSlidesFromBase64`.

### 2. MIAI Bridge (`packages/bridge`)
Lokal service på `https://127.0.0.1:<port>`.

| Endpoint | Formål |
|----------|--------|
| `GET  /health` | Status, aktiv provider/model, hardware-info |
| `GET  /models` | Tilgængelige modeller på tværs af providere |
| `POST /v1/tasks/{taskId}` | Kør en navngiven opgave (fx `grammar`, `explain-formula`) med kontekst; SSE-stream |
| `POST /v1/chat` | Fri chat med dokumentkontekst og evt. tools; SSE-stream |
| `GET/PUT /settings` | Brugerindstillinger |
| `POST /v1/rag/*` | Indeksering og søgning (fase 4) |

Ansvar:
- **Provider-router**: `OllamaProvider`, `OpenAICompatibleProvider` (LM Studio,
  llama.cpp, vLLM, Foundry Local), `AnthropicProvider`, `AzureOpenAIProvider`,
  `BedrockProvider`. Fælles interface: `chat(messages, tools?, stream)`.
- **Opgavekatalog**: hver opgave = prompt-skabelon + output-skema + anbefalet model.
  Skabeloner ligger som filer (`tasks/*.yaml`) og kan overskrives af brugeren.
- **Strukturerede svar**: JSON-skema (fx `{ edits: [{ original, replacement, reason }] }`
  til grammatik) håndhæves via provider-understøttelse (Ollama `format`,
  llama.cpp grammar) eller validering + retry.
- **Sikkerhed**: bind kun til loopback, token udstedt ved installation og gemt i
  add-in'ets `OfficeRuntime.storage`, CORS-whitelist.

### 3. Model-runtime (eksternt)
Ollama som anbefalet standard (nem installation, modelstyring, OpenAI-kompatibel
API). Bridgen kan tilbyde at installere/hente den anbefalede model.

## Hvorfor en bridge og ikke direkte kald fra add-in?

1. Add-ins skal serveres over HTTPS; kald til `http://localhost:11434` giver mixed
   content og CORS-problemer (Ollama kræver `OLLAMA_ORIGINS`).
2. Chromium-baserede webviews indfører Private/Local Network Access-restriktioner.
3. Provider-nøgler (cloud) må ikke ligge i en webview.
4. Filsystemadgang (RAG, skabeloner, logs) kræver en native proces.
5. Ét sted for logging/audit af hvad der sendes hvorhen.

## Dataflow – eksempel "Forklar formel" (Excel)

1. Bruger markerer `C5` og klikker *Forklar formel*.
2. `ExcelAdapter.getSelection()` → formel, værdier i refererede celler (begrænset), kolonneoverskrifter.
3. `POST /v1/tasks/explain-formula` → bridge vælger model, bygger prompt, streamer svar.
4. Task pane viser forklaring; valgfrit *Indsæt som kommentar* på cellen.

## Repo-struktur (planlagt)

```
MIAIPlugin/
├─ packages/
│  ├─ addin/        # Office.js task pane + manifest
│  ├─ bridge/       # lokal service
│  └─ shared/       # typer, opgave-skemaer, prompt-skabeloner
├─ tasks/           # prompt-skabeloner (YAML) pr. opgave
├─ evals/           # eval-sæt og harness
├─ installer/       # Windows (MSI/winget) og macOS (pkg)
└─ docs/            # plan, arkitektur, ADR'er
```
