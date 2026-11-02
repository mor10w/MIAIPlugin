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
Office (task pane)  ──https://localhost:44321──▶  MIAI Bridge  ──OpenAI API──▶  dit inference-værktøj
```
Den lokale **bridge** serverer task pane, holder prompts og indstillinger og håndhæver en **politik** for, hvilke hosts der må kaldes. Som standard er det kun loopback. Det gør det muligt senere at køre i organisationer, der peger på en intern inference-server.

## Kom i gang
Se [docs/UDVIKLING.md](docs/UDVIKLING.md). Kort:
```bash
npm install && npm run certs && npm run build && npm start
```
Sideload derefter manifestet med `npm run sideload:mac` eller `npm run sideload:windows`.

## Dokumentation
- [Projektplan](docs/PLAN.md)
- [Arkitektur](docs/ARKITEKTUR.md)
- [Udvikling og test](docs/UDVIKLING.md)
- [Arkitekturbeslutninger (ADR)](docs/adr/)

## Struktur
```
packages/shared   typer, opgavekatalog og prompts
packages/bridge   lokal Node.js-service (Fastify)
packages/addin    Office.js task pane (React + Fluent UI) og manifest
```
