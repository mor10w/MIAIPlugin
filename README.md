# MIAIPlugin

AI-assistent til Microsoft **Word, Excel og PowerPoint**, der bruger **dine egne modeller**. Den kan tale med ethvert lokalt inference-værktøj med OpenAI-kompatibelt API, fx Ollama, LM Studio, llama.cpp, vLLM, LocalAI eller Jan. Dokumenter forlader ikke maskinen.

> Status: fase 0 (spike). Fungerende skelet; ikke afprøvet i Office endnu, se [testlisten](docs/UDVIKLING.md#fase-0--manuel-testliste).

## Funktioner
| | Word | Excel | PowerPoint |
|---|:-:|:-:|:-:|
| Ret stavning og grammatik (som sporede ændringer i Word) | ✅ | | ✅ |
| Omskriv i anden tone · parafrasér · forenkl | ✅ | | ✅ |
| Oversæt | ✅ | ✅ | ✅ |
| Opsummér | ✅ | | ✅ |
| Skriv artikel ud fra én linje | ✅ | | |
| Præsentationsdisposition | ✅ | | ✅ |
| Tekst til slides | | | ✅ |
| Forklar formel · lav formel ud fra beskrivelse | | ✅ | |
| Chat med dokumentet | ✅ | ✅ | ✅ |

## Sådan virker det
```
Office (task pane) ──OpenAI API──────────────────────────────▶ dit inference-værktøj     (direkte, standard)
Office (task pane) ──https://localhost:44321──▶ MIAI Bridge ──▶ dit inference-værktøj     (valgfri bridge)
```
- **Direkte**: add-in'et hostes statisk (GitHub Pages) og kalder selv modellen. Du skal kun bruge inference-værktøjet og ét manifest, og værktøjet skal tillade CORS fra add-in'ets adresse.
- **Bridge** (valgfri): en lokal service serverer add-in'et og kalder modellen. Den giver låst politik for organisationer, audit-log og UI uden internet.

Begge tilstande bruger samme motor og samme **politik** for, hvilke hosts der må kaldes (standard: kun loopback). Tilstanden vælges automatisk. Se [ADR-0007](docs/adr/0007-bridge-valgfri-direkte-forbindelse.md).

## Kom i gang
Se [docs/UDVIKLING.md](docs/UDVIKLING.md). Kort, til udvikling med direkte forbindelse:
```bash
npm install && npm run certs && npm run dev
```
Sideload derefter `packages/addin/manifest.dev.xml` med `npm run sideload:mac` eller `npm run sideload:windows`.

## Dokumentation
- [Projektplan](docs/PLAN.md)
- [Arkitektur](docs/ARKITEKTUR.md)
- [Udvikling og test](docs/UDVIKLING.md)
- [Arkitekturbeslutninger (ADR)](docs/adr/)

## Struktur
```
packages/shared   fælles motor: opgaver/prompts, OpenAI-klient, politik
packages/addin    Office.js task pane (React + Fluent UI), backends, manifest-skabelon
packages/bridge   valgfri lokal Node.js-service (Fastify)
```
