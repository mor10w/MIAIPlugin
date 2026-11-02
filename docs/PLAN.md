# MIAIPlugin – projektplan (udkast v0.1)

> Status: **Udkast** · Dato: 2026-11-02 · Ejer: M. Wiingaard

## 1. Vision

En AI-assistent til Microsoft Word, Excel, PowerPoint (og senere Outlook), der
dækker det samme funktionsområde som kommercielle "AI writing assistants" og
Claude-integrationer, men hvor **al inferens som standard sker på brugerens egen
maskine**. Ingen dokumentdata forlader maskinen, medmindre brugeren bevidst
vælger en cloud-udbyder.

**Kerneløfter**

1. **Privat som standard**: lokale modeller (Ollama, LM Studio, llama.cpp, vLLM,
   Foundry Local) via en OpenAI-kompatibel API. Ingen telemetri.
2. **Leverandøruafhængig**: én provider-abstraktion. Cloud-udbydere (Claude,
   Azure OpenAI, AWS Bedrock) kan slås til, men kun ved aktivt valg.
3. **Native i Office**: arbejder på markeringen, viser forslag som diff,
   track changes, kommentarer eller nye slides, og kan altid fortrydes.
4. **Dansk + engelsk** UI og prompts fra dag ét.

## 2. Funktionskatalog (målbillede)

| # | Funktion | Word | Excel | PowerPoint | Outlook | Fase |
|---|----------|:----:|:-----:|:----------:|:-------:|:----:|
| F1 | Ret stavning og grammatik (som track changes) | ✅ | – | ✅ | ✅ | MVP |
| F2 | Omskriv / parafrasér / forenkl / skift tone | ✅ | – | ✅ | ✅ | MVP |
| F3 | Oversæt markering eller hele dokumentet | ✅ | ✅ (celler) | ✅ | ✅ | MVP |
| F4 | Fri chat om dokumentet ("spørg dit dokument") | ✅ | ✅ | ✅ | ✅ | MVP |
| F5 | Generér fuld artikel/blogindlæg ud fra én linje eller disposition | ✅ | – | – | – | 2 |
| F6 | Salgstekst / persuasive copy ud fra skabelon | ✅ | – | ✅ | ✅ | 2 |
| F7 | Forklar formel i almindeligt sprog | – | ✅ | – | – | MVP |
| F8 | Generér formel ud fra beskrivelse (inkl. validering) | – | ✅ | – | – | 2 |
| F9 | Generér præsentationsdisposition | ✅ | – | ✅ | – | 2 |
| F10 | Tekst → slides (fra Word-markering eller fri tekst) | – | – | ✅ | – | 2 |
| F11 | Opsummér dokument / tråd | ✅ | – | ✅ | ✅ | 2 |
| F12 | Dataanalyse: beskriv tabel, foreslå pivot/diagram | – | ✅ | – | – | 3 |
| F13 | Agent-tilstand: modellen bruger værktøjer (læs/skriv range, indsæt slide …) | ✅ | ✅ | ✅ | – | 3 |
| F14 | Brugerdefinerede prompts / "skills" (delt via fil eller repo) | ✅ | ✅ | ✅ | ✅ | 3 |
| F15 | Lokal RAG over egne dokumenter (mapper på disken) | ✅ | ✅ | ✅ | ✅ | 4 |

F13 er det, der gør produktet mere end en prompt-knap og nærmer sig
"Claude i Office": modellen får tool-calling-adgang til en lille, sikker API over
Office-dokumentet og kan selv planlægge flere trin.

## 3. Arkitektur (resumé)

Detaljer i [ARKITEKTUR.md](ARKITEKTUR.md).

```
┌──────────────── Office (Word / Excel / PowerPoint / Outlook) ────────────────┐
│  Task pane (React + TypeScript + Fluent UI)   Ribbon-knapper / kontekstmenu   │
│        │  Office.js host-adaptere (WordAdapter, ExcelAdapter, PptAdapter)     │
└────────┼──────────────────────────────────────────────────────────────────────┘
         │ HTTPS (localhost, selvsigneret cert installeret i brugerens store)
┌────────▼──────────────── MIAI Bridge (lokal companion-service) ──────────────┐
│  Provider-router · prompt-skabeloner · streaming (SSE) · lokal RAG · log      │
└────────┬─────────────────────────┬───────────────────────────┬────────────────┘
         │                         │                           │ (valgfrit, opt-in)
   Ollama / LM Studio        llama.cpp / vLLM            Claude / Azure OpenAI / Bedrock
```

**Hovedbeslutninger (foreløbige, se `docs/adr/`)**

- **Office Web Add-in (Office.js)** frem for VSTO/COM: virker på Windows, Mac og
  Office på web, distribueres via manifest, og er fremtidssikret. (ADR-0001)
- **Lokal bridge-service** mellem add-in og model: løser HTTPS/mixed content,
  CORS og Private Network Access, samler provider-logik ét sted og giver adgang
  til filsystem (RAG). (ADR-0002)
- **OpenAI-kompatibel API som fællesnævner** mod lokale runtimes; native adapter
  til Anthropic Messages API for Claude. (ADR-0003)

## 4. Modelstrategi (lokalt)

| Brug | Forslag (evalueres i fase 0) | Krav |
|------|------------------------------|------|
| Standard skrive/omskrive, dansk | Gemma 3 12B / Qwen3 14B / Mistral Small 3.x (Q4_K_M) | ~10–16 GB RAM/VRAM |
| Lille/hurtig (laptop uden GPU) | Qwen3 4B / Gemma 3 4B / Phi-4-mini | ~4–6 GB |
| Formler + tool calling | Qwen3 (tool calling stabil) / Qwen2.5-Coder | function calling |
| Oversættelse | Samme som standard; evt. dedikeret model | – |
| Embeddings (RAG) | bge-m3 / nomic-embed-text / multilingual-e5 | flersproget |

Modelvalg er **konfiguration, ikke kode**. Fase 0 laver et lille eval-sæt på dansk
og engelsk (grammatik, omskrivning, formelforklaring, slides) så valget kan
genbedømmes, når nye modeller kommer.

Hardware-detektion i bridgen foreslår en model ud fra RAM/VRAM ved første opstart.

## 5. Faseplan

### Fase 0 – Afklaring og spike (1–2 uger)
- [ ] Spike: Office.js task pane → `https://localhost` bridge → Ollama, med streaming i Word desktop **og** Office på web.
- [ ] Verificér Private Network Access / CORS-adfærd i WebView2 (Windows), WKWebView (Mac) og browser.
- [ ] Verificér PowerPoint API-niveau til tekst → slides (shapes/tekstbokse vs. `insertSlidesFromBase64`).
- [ ] Eval-sæt (30–50 opgaver) + valg af 2–3 standardmodeller.
- [ ] Beslut licens, distributionsform og navn (se §8).

### Fase 1 – MVP (4–6 uger)
- Monorepo, CI, manifest (unified JSON manifest + XML-fallback).
- Bridge: provider-abstraktion, Ollama + OpenAI-kompatibel provider, SSE-streaming, indstillinger, sundhedstjek.
- Task pane: chat, hurtighandlinger (F1–F4, F7), model-vælger, "Indsæt / Erstat / Som kommentar / Som track changes".
- Word- og Excel-adaptere; PowerPoint kun tekst i markeret shape.
- Installer til Windows (MSI/winget) der installerer bridge, cert og sideloader manifest.

### Fase 2 – Indholdsgenerering (4–6 uger)
- F5, F6, F8–F11. Tekst → slides med layout fra aktivt master/tema.
- Prompt-skabeloner som versionerede filer (YAML/Markdown), redigerbare af brugeren.
- Mac-installer (pkg/Homebrew).

### Fase 3 – Agent og udvidelse (6+ uger)
- F12–F14: tool calling-loop med whitelisted Office-værktøjer, plan → bekræft → udfør, fuld fortryd.
- Valgfri cloud-providere (Claude, Azure OpenAI, Bedrock) bag eksplicit opt-in og tydelig markering i UI.
- Outlook-add-in.

### Fase 4 – Lokal viden (RAG)
- F15: indeksér valgte mapper (docx, pdf, xlsx, md) i lokal vektorstore (SQLite + sqlite-vec / LanceDB).
- Kildehenvisninger i svar.

## 6. Ikke-funktionelle krav

- **Privatliv**: ingen netværkstrafik ud af maskinen i lokal tilstand; kontrollérbart (bridge logger alle udgående kald). Ingen telemetri.
- **Sikkerhed**: bridgen lytter kun på `127.0.0.1`, kræver pr.-installation token, CORS låst til add-in-origin. Dokumentindhold behandles som data i prompts (prompt-injection-bevidsthed i agent-tilstand: ingen destruktive værktøjer uden bekræftelse).
- **Ydelse**: første token < 2 s på anbefalet hardware for korte opgaver; streaming altid.
- **Robusthed**: store dokumenter chunkes; aldrig stille overskrivning – alle ændringer kan fortrydes.
- **Tilgængelighed**: Fluent UI, tastaturnavigation, høj kontrast.
- **Enterprise**: central udrulning via Microsoft 365 Admin Center + konfig via GPO/Intune (fase 3).

## 7. Teknologistak

| Lag | Valg | Alternativ |
|-----|------|-----------|
| Add-in | TypeScript, React, Fluent UI v9, Vite, Office.js | – |
| Bridge | Node.js (TypeScript, Fastify) **eller** .NET 8 (ASP.NET Minimal API) | Rust (Tauri) |
| Model-runtime | Ollama (default), LM Studio, llama.cpp-server, Foundry Local | vLLM (server/GPU) |
| RAG | SQLite + sqlite-vec | LanceDB |
| Test | Vitest, Playwright, Office add-in debugging, eval-harness (promptfoo) | – |
| CI | GitHub Actions (lint, test, build, signering) | Azure DevOps |

Bridge-sprog er en åben beslutning (ADR-0004): Node giver ét sprog i hele
stakken; .NET giver nemmere Windows-service, cert-håndtering og MSI.

## 8. Åbne spørgsmål

1. Målgruppe: privat/freelance, SMB eller enterprise med central udrulning?
2. Licens: open source (MIT/Apache-2.0) eller kommerciel/dual?
3. Distribution: AppSource kræver HTTPS-hostet add-in – hostes UI fra bridgen lokalt eller fra statisk CDN? (påvirker offline-krav)
4. Skal Office på web være understøttet i MVP, eller kun desktop?
5. Navn: "MIAIPlugin" er arbejdstitel. Undgå navne/branding der ligner eksisterende produkter.
6. Bridge-sprog: Node.js eller .NET?

## 9. Risici

| Risiko | Konsekvens | Afbødning |
|--------|-----------|-----------|
| Browser/WebView blokerer kald til localhost (PNA, mixed content) | Add-in kan ikke nå model | HTTPS-bridge med lokalt cert; spike i fase 0 |
| Lokale modeller er svagere på dansk end cloud | Lavere kvalitet | Eval-sæt, modelvalg som konfiguration, valgfri cloud |
| Svag hardware hos brugeren | Langsom oplevelse | Små modeller, hardware-detektion, tydelig forventningsstyring |
| Office.js API-huller (især PowerPoint) | Funktioner begrænses | Tjek requirement sets; fallback via base64-pptx genereret i bridgen |
| Prompt injection via dokumentindhold i agent-tilstand | Uønskede ændringer | Whitelist, bekræftelse, fortryd, ingen netværksværktøjer |
