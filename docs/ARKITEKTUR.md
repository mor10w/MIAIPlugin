# Arkitektur (udkast)

## Komponenter

### 1. Office Add-in (`packages/addin`)
- Office Web Add-in med task pane + ribbon-kommandoer (function commands).
- **Host-adaptere** skjuler forskelle mellem Word, Excel og PowerPoint bag ét interface:

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

- Ændringer vises som diff i task pane før de anvendes. Word: `changeTrackingMode`
  slås midlertidigt til, så brugeren kan acceptere/afvise ændringer i Word.
- Word: markeringen gemmes som sporet objekt (`trackedObjects`), så resultatet
  indsættes det rigtige sted, selvom brugeren klikker andetsteds imens.
- Excel: formlen skrives i den markerede celle, hvorefter resultatet læses. Giver
  Excel en fejl (#NAME? osv.), får brugeren besked. De tidligere formler gemmes,
  så **Fortryd** virker (Ctrl+Z fortryder ikke altid add-in-ændringer).
- PowerPoint: tekst → slides via `slides.add`. Layoutets pladsholdere udfyldes,
  og findes de ikke, tilføjes tekstbokse (PowerPointApi 1.4). Hvis det viser sig
  utilstrækkeligt med firmaskabeloner, er plan B en `.pptx` genereret i bridgen
  og indsat med `insertSlidesFromBase64`.

### 2. MIAI Bridge (`packages/bridge`)
Lokal Node.js-service (Fastify) på `https://127.0.0.1:44321`. Serverer både task pane (statiske filer) og API'et, så de deler origin.

| Endpoint | Formål |
|----------|--------|
| `GET  /api/health` | Status og om den aktive endpoint svarer |
| `GET  /api/models` | Modeller fra endpointens `/v1/models` |
| `GET  /api/tasks` | Opgavekatalog |
| `POST /api/tasks/{id}` | Kør opgave (`grammar`, `rewrite`, `translate`, `explain-formula`, `text-to-slides`, `chat` …); svar som SSE |
| `GET/PUT /api/settings` | Indstillinger og politik (API-nøgler maskeres) |

Ansvar:
- **Én OpenAI-klient** (`openai.ts`): `/v1/models` og streaming `/v1/chat/completions`. Tåler servere, der ignorerer `stream: true`.
- **Opgavekatalog** (`@miai/shared/tasks.ts`): prompt, output-type (tekst eller JSON med Zod-skema), parametre og indsætningsmåder.
- **Struktureret output**: `response_format` efter endpointens `structuredOutput` (`json_schema`/`json_object`/`none`). Skemaet sendes også i prompten, og svaret udtrækkes robust, valideres med Zod og prøves én gang mere ved fejl.
- **`<think>`-filter** for reasoning-modeller, også når tags er splittet over chunks.
- **Politik** (`allowedHosts`), kontrolleret både ved gem og ved hvert kald (ADR-0006).
- **Sikkerhed**: lytter kun på loopback, Host-header-tjek (DNS rebinding), `X-MIAI`-header og origin-tjek, så fremmede websider ikke kan bruge bridgen. Modeloutput renses med DOMPurify før visning og indsættelse.
- **Audit-log** (valgfri): tid, opgave, host, model, varighed og tegn/tokens, men aldrig indhold.

### 3. Inference-værktøj (eksternt)
Et hvilket som helst OpenAI-kompatibelt værktøj (Ollama som standardforslag). Modeller hentes og styres i værktøjet selv.

## Hvorfor en bridge og ikke direkte kald fra add-in?

1. Add-ins skal serveres over HTTPS; kald til `http://localhost:11434` giver mixed
   content og CORS-problemer (Ollama kræver `OLLAMA_ORIGINS`).
2. Chromium-baserede webviews indfører Private/Local Network Access-restriktioner.
3. API-nøgler (fx til en intern gateway) må ikke ligge i en webview.
4. Filsystemadgang (RAG, skabeloner, logs) kræver en native proces.
5. Ét sted for logging/audit af hvad der sendes hvorhen.

## Dataflow – eksempel "Forklar formel" (Excel)

1. Bruger markerer `C5` og klikker *Forklar formel*.
2. `ExcelAdapter.getContext()` → formel, værdier i refererede celler (begrænset), kolonneoverskrifter.
3. `POST /api/tasks/explain-formula` → bridge vælger model, bygger prompt, streamer svar.
4. Task pane viser forklaring; valgfrit *Indsæt som kommentar* på cellen.

## Repo-struktur

```
MIAIPlugin/
├─ packages/
│  ├─ addin/        # Office.js task pane + manifest
│  ├─ bridge/       # lokal service
│  └─ shared/       # typer, opgavekatalog og prompts, hjælpefunktioner
├─ scripts/         # ikoner, sideload (Mac/Windows)
├─ evals/           # (fase 0) eval-sæt og harness
├─ installer/       # (fase 1) Windows (MSI/winget) og macOS (pkg)
└─ docs/            # plan, arkitektur, udvikling, ADR'er
```
