# ADR-0007: Bridgen er valgfri – direkte forbindelse er standard

- Status: Vedtaget
- Dato: 2026-11-02
- Erstatter: [ADR-0002](0002-lokal-bridge-service.md)

## Kontekst
ADR-0002 gjorde en lokal bridge-service obligatorisk. Begrundelsen holdt kun delvist:

- **Mixed content** er ikke et problem. Chromium (WebView2) regner `http://127.0.0.1` og `http://localhost` som sikre, så en HTTPS-side må kalde dem. Det er verificeret med Playwright/Chromium (se nedenfor). WebKit på Mac skal stadig afprøves.
- **CORS** er reelt, men løses med en indstilling i inference-værktøjet:
  - Ollama tillader `https://localhost:*` som standard; andre origins kræver `OLLAMA_ORIGINS`.
  - LM Studio har en "Enable CORS"-indstilling.
  - vLLM har `--allowed-origins`.
- For privatpersoner og SMB er en ekstra service med eget certifikat en stor installationsbyrde.

## Beslutning
Al opgavelogik (OpenAI-klient, prompts, `<think>`-filter, JSON-validering, politik) ligger i `@miai/shared` og kører både i browseren og i Node. Add-in'et har to backends med samme interface (`packages/addin/src/backend`):

| | **Direkte** (standard) | **Bridge** (valgfri) |
|---|---|---|
| Hvem kalder modellen | Add-in'et selv | MIAI Bridge |
| Add-in'ets filer | Hostes statisk (GitHub Pages, eget domæne, intern webserver) | Serveres af bridgen fra `https://localhost:44321` |
| Indstillinger | `localStorage` i add-in'et | Konfigurationsfil |
| Politik | Standard loopback; kan ændres af brugeren | Maskinpolitik kan låses af organisationen |
| Kræver | Manifest + CORS i inference-værktøjet | Installation af service og certifikat |
| Ekstra | – | Offline UI, audit-log, senere RAG over lokale filer |

Tilstanden vælges automatisk ved opstart. Svarer `/api/health` på samme origin som en MIAI Bridge, bruges bridgen. Ellers bruges direkte forbindelse. Det er samme build. Kun manifestet afgør, hvor add-in'et hentes fra:

- `manifest.xml`: bridge (`https://localhost:44321`)
- `manifest.dev.xml`: Vite-udviklingsserver (`https://localhost:3000`), direkte
- hostet manifest: genereres af Pages-workflowet til `<pages-url>/manifest.xml`, direkte

## Verifikation (Chromium/Playwright, 2026-11-02)
- Direkte, add-in på `https://localhost:3000`, mock-model på `http://127.0.0.1:11434` med CORS: forbundet, streaming, indsættelse i (simuleret) Word.
- Samme uden CORS-header: blokeret af browseren. UI viser "Ingen forbindelse", og fejlteksten forklarer CORS-opsætning.
- Bridge-tilstand: samme flow uden CORS.

## Konsekvenser
+ Installation for privat og SMB: inference-værktøj + ét manifest.
+ Bridgen bevares til organisationer (låst politik, audit, intern hosting) og senere funktioner.
- Direkte tilstand kræver CORS-opsætning for nogle værktøjer og for hostede origins.
- Hostet add-in kræver internet for at hente UI'et (dokumentdata går stadig kun til den lokale model), og brugeren stoler på den hostede kode.
- Nyere Chromium indfører "Local Network Access"-tilladelse, når offentlige sites kalder localhost. Det skal afprøves i WebView2 med et hostet (ikke-localhost) origin. Fallback er bridgen eller hosting på `https://localhost`.
