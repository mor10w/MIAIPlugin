# MIAIPlugin – projektplan

> Status: **Udkast v0.3** · Senest opdateret: 2026-11-02 · Ejer: M. Wiingaard

## 1. Vision

En AI-assistent til Microsoft Word, Excel, PowerPoint og Outlook. Den
dækker samme funktioner som kommercielle "AI writing assistants", men **al
inferens sker på en OpenAI-kompatibel endpoint, som brugeren selv kontrollerer**.
Som standard er det et lokalt inference-værktøj på brugerens egen maskine.

**Kerneløfter**

1. **Ét API: OpenAI Chat Completions.** Alt, der taler `/v1/chat/completions` og
   `/v1/models`, kan bruges: Ollama, LM Studio, llama.cpp-server, vLLM, LocalAI,
   Jan, KoboldCpp, text-generation-webui, Foundry Local, TGI m.fl. Ingen
   leverandørspecifikke adaptere. (ADR-0003)
2. **Data forbliver hos brugeren.** Som standard er kun endpoints på loopback
   (`127.0.0.1`, `localhost`, `::1`) tilladt. Ingen telemetri. (ADR-0006)
3. **Native i Office:** arbejder på markeringen, viser resultatet før det
   anvendes, kan indsætte som track changes eller kommentar og kan fortrydes.
4. **Dansk og engelsk** i UI og prompts.

## 2. Målgruppe og scope

| | Fase 1–2 (nu) | Senere (holdes åbent i designet) |
|---|---|---|
| Målgruppe | Privatpersoner, freelancere, mindre virksomheder | Større organisationer uden ønske om 3rd party cloud |
| Office | **Lokale desktop-installationer** (Windows, Mac) | Office på web, central udrulning (M365 Admin Center) |
| Inference | Lokal på brugerens maskine | On-prem inference-server på intern host (vLLM/TGI bag gateway) |
| Konfiguration | Brugerens egen `config.json` via UI | Maskin-/politikfil (GPO/Intune) der låser endpoints og modeller |
| Installation | Inference-værktøj + ét manifest (add-in hostet statisk, direkte forbindelse) | MIAI Bridge pakket til Intune/SCCM, Jamf |

**Designregler der holder enterprise-døren åben** (koster næsten intet nu):
- Endpoint = `baseUrl` + valgfri API-nøgle. En intern `https://llm.firma.local/v1` er bare en anden endpoint.
- Alle udgående adresser valideres mod en **politik** (`allowedHosts`). Politikken kan komme fra en maskinfil, som brugeren ikke kan overskrive. (ADR-0006)
- Ingen funktion må kræve internet ud over Microsofts Office.js-CDN.
- Audit-log over alle udgående kald (host, model, tokens, ikke indhold) kan slås til.

## 3. Funktionskatalog

| # | Funktion | Word | Excel | PowerPoint | Fase |
|---|----------|:----:|:-----:|:----------:|:----:|
| F1 | Ret stavning og grammatik (som track changes) | ✅ | – | ✅ | 1 |
| F2 | Omskriv / parafrasér / forenkl / skift tone | ✅ | – | ✅ | 1 |
| F3 | Oversæt markering | ✅ | ✅ | ✅ | 1 |
| F4 | Chat om dokumentet | ✅ | ✅ | ✅ | 1 |
| F7 | Forklar formel i almindeligt sprog | – | ✅ | – | 1 |
| F8 | Generér formel ud fra beskrivelse (med validering og fortryd) | – | ✅ | – | 1 |
| F5 | Generér artikel/blogindlæg ud fra én linje | ✅ | – | – | 1 |
| F9 | Generér præsentationsdisposition | ✅ | – | ✅ | 1 |
| F10 | Tekst → slides | – | – | ✅ | 1 |
| F11 | Opsummér | ✅ | – | ✅ | 1 |
| F6 | Salgstekst / persuasive copy | ✅ | – | ✅ | 2 |
| F12 | Dataanalyse: beskriv tabel, foreslå pivot/diagram | – | ✅ | – | 2 |
| F14 | Brugerdefinerede prompts / opgaver (filer) | ✅ | ✅ | ✅ | 2 |
| F13 | Agent-tilstand med tool calling på dokumentet | ✅ | ✅ | ✅ | 3 |
| F15 | Lokal RAG over egne mapper | ✅ | ✅ | ✅ | 4 |
| –  | Outlook | | | | 3 |

## 4. Arkitektur (resumé)

Detaljer i [ARKITEKTUR.md](ARKITEKTUR.md). Bridgen er **valgfri** (ADR-0007).

```
┌──────────── Office desktop (Word / Excel / PowerPoint) ────────────┐
│ Task pane (React + Fluent UI) · host-adaptere (Office.js)           │
│ @miai/shared: opgaver/prompts · OpenAI-klient · <think>-filter ·    │
│               JSON-validering · politik                             │
│ Backend:  [Direkte]  ─────────────┐          [Bridge] ──┐           │
└───────────────────────────────────┼─────────────────────┼───────────┘
                                    │ OpenAI API (CORS)   │ https://localhost:44321/api
                                    │                     ▼
                                    │       MIAI Bridge (valgfri): samme motor,
                                    │       låst politik, audit-log, offline UI
                                    ▼                     │
             Ollama · LM Studio · llama.cpp · vLLM · LocalAI · Jan · … ◀┘
```

- **Direkte** (standard for privat/SMB): add-in'et hostes statisk (GitHub Pages) og kalder selv modellen.
- **Bridge** (organisationer, offline, senere RAG): add-in'et serveres af bridgen, som kalder modellen.

## 5. Modelstrategi

Modelvalg er konfiguration. Bridgen lister modeller fra endpointen
(`/v1/models`); brugeren vælger en standardmodel og evt. en model pr. opgave.

| Brug | Forslag (evalueres i fase 0) | Ca. krav (Q4) |
|------|------------------------------|-------------|
| Standard, dansk | Gemma 3 12B, Qwen3 14B, Mistral Small 3.x | 10–16 GB |
| Laptop uden GPU | Qwen3 4B, Gemma 3 4B, Phi-4-mini | 4–6 GB |
| Formler, struktureret output | Qwen3, Qwen2.5-Coder | – |

Forslagene evalueres med eval-sættet i `packages/evals` (se [EVAL.md](EVAL.md)).

**Embedding-modeller (F15, lokal RAG – ADR-0013):** embeddings hentes fra samme endpoint via `/v1/embeddings`.

| Brug | Model | Begrundelse |
|------|-------|-------------|
| Standard (dansk og blandet) | **bge-m3** (568M, 1.024 dim.) | Trænet på 100+ sprog inkl. dansk, lang kontekst (8k), ingen præfikser |
| Laptop uden GPU | embeddinggemma (308M, 768 dim.) | Flersproget og lille; MIAI tilføjer modellens præfikser |
| Engelsk, hurtigst | nomic-embed-text (137M, 768 dim.) | Meget udbredt i Ollama; svagere på dansk |
| GPU-server | qwen3-embedding (0,6B–8B) | Stærk flersproget søgning til store mængder |

Valget er baseret på modellernes dokumenterede egenskaber. Det skal bekræftes med eval-sættet på dansk materiale, når der køres på rigtig hardware (#2).

Struktureret output (grammatik-rettelser, slides, formler) bruger
`response_format: json_schema`, når endpointen understøtter det. Ellers bruges
`json_object` eller prompt + parsing + ét genforsøg. Dette sættes pr. endpoint
(`structuredOutput`).

Reasoning-modeller: `<think>…</think>` og `reasoning_content` filtreres fra af den fælles motor.

## 6. Faseplan

### Fase 0 – Spike (påbegyndt)
- [x] Monorepo-skelet: `shared`, `bridge`, `addin`.
- [x] Bridge: HTTPS, statiske add-in-filer, `/api/health`, `/api/models`, `/api/tasks/:id` (SSE), `/api/settings`, politik-check.
- [x] Add-in: task pane, host-adaptere, hurtighandlinger, chat, indstillinger.
- [x] Bridgen gjort valgfri: fælles motor i `shared`, direkte backend i add-in'et, automatisk valg af tilstand (ADR-0007).
- [x] Browsertest (Chromium) af direkte tilstand og bridge-tilstand med simuleret Word.
- [x] GitHub Pages-hosting + genererede manifester (bridge, dev, hostet).
- [ ] Direkte tilstand i WebView2 med hostet origin (Local Network Access-prompt?) og i WKWebView på Mac.
- [ ] Verificér end-to-end i Word, Excel og PowerPoint desktop på Windows og Mac (manuel test).
- [ ] Eval-sæt (30–50 opgaver, dansk/engelsk) og valg af anbefalede modeller.

### Fase 1 – MVP
- Hærdning af F1–F11, fejlhåndtering, tokengrænser og chunking af store dokumenter.
- Førstegangsguide i direkte tilstand: find kørende værktøjer på kendte porte og vis præcis CORS-opsætning.
- Distribution: hostet manifest (direkte) som standard. Valgfri installer til bridgen (Windows MSI/winget, macOS pkg) med lokalt CA/cert og service.

### Fase 2 (færdig 2026-11-03)
- [x] F6, F12: Salgstekst (word+powerpoint) og dataanalyse (excel) som nye opgaver.
- [x] F14: Opgaver som brugerredigerbare JSON-filer, deling via mappe (ADR-0011); "Opgaver"-fanen med redigering, import/eksport; maskin-opgaver via `MIAI_TASKS_DIR`.
- [x] Politikfil og audit-log færdiggjort (forberedelse til enterprise): JSONL `audit.log` med rotation i bridgen, ringbuffer i direkte tilstand; log vises under Indstillinger.

### Fase 3
- [x] F13: agent-tilstand (whitelisted Office-handlinger, plan → bekræft → udfør, fortryd af hele sessionen) – ADR-0012.
- [x] Distribution (#5): MIAI App til Windows (NSIS pr. bruger) og Mac (arm64/x64), automatisk opdatering via electron-updater, release-workflow med signering og notarisering, når certifikater er sat op, og winget-skabelon – ADR-0016. Mangler: køb af certifikater (Apple Developer ID, Azure Trusted Signing).
- [x] Outlook (læse- og skrivetilstand, Svar på mail) og Office på web via central udrulning med `miai-config.json` og intern inference-server – ADR-0015.

### Fase 4
- [x] F15: lokal RAG (node:sqlite + vektorsøgning i bridgen, embeddings via `/v1/embeddings` på samme endpoint, kilder i chatten) – ADR-0013.

## 7. Ikke-funktionelle krav

- **Privatliv**: motoren kalder kun hosts, som politikken tillader (i både direkte tilstand og bridge); ingen telemetri.
- **Sikkerhed**: modeloutput renses (DOMPurify) før visning og indsættelse, og dokumentindhold markeres som data i prompts. Bridgen lytter kun på loopback, kræver korrekt `Host`-header (DNS rebinding) og eget `Origin`. I direkte tilstand afgør inference-værktøjets CORS-opsætning, hvilke origins der må kalde det.
- **Ydelse**: streaming altid for tekstsvar; første token < 2 s på anbefalet hardware.
- **Fortryd**: Word bruger track changes eller Ctrl+Z. Excel gemmer de tidligere formler og har en "Fortryd"-knap, fordi add-in-ændringer ikke altid kan fortrydes med Ctrl+Z.

## 8. Teknologistak

| Lag | Valg |
|-----|------|
| Add-in | TypeScript, React 19, Fluent UI v9, Vite, Office.js, XML-manifest |
| Bridge | Node.js 22 LTS, TypeScript, Fastify 5, Zod (ADR-0004) |
| Test | Vitest (shared og bridge mod mock OpenAI-server); Playwright-browsertest af task pane; senere eval-harness |
| CI | GitHub Actions |

## 9. Beslutninger og åbne spørgsmål

Besluttet:
- Kun OpenAI-kompatibel API (ADR-0003).
- Desktop først; privat og SMB først (ADR-0005).
- Bridge i Node.js/TypeScript (ADR-0004).
- Politik/allowedHosts fra start (ADR-0006).
- Bridgen er valgfri; direkte forbindelse er standard (ADR-0007).
- Licens: MIT License – koden er åben for fri brug; ophavsret er ManInfo ApS.

Åbne:
1. Endeligt produktnavn.
2. Kodesignering af installer: certifikat og udgiver.

## 10. Risici

| Risiko | Afbødning |
|--------|-----------|
| CORS ikke sat op i inference-værktøjet (direkte tilstand) | Tydelig fejltekst med præcis opsætning; førstegangsguide; bridge som alternativ |
| Local Network Access-prompt i WebView2 ved hostet origin | Afprøves i fase 0; alternativt hosting på `https://localhost` eller bridge |
| Office-webview stoler ikke på lokalt cert (bridge/dev) | Installer tilføjer CA til brugerens trust store; i udvikling bruges `office-addin-dev-certs` |
| Lokale modeller svagere på dansk | Eval-sæt; model pr. opgave; enterprise kan pege på større on-prem model |
| Svag hardware | Små modeller, anbefaling ud fra RAM |
| Endpoints understøtter struktureret output forskelligt | `structuredOutput`-indstilling pr. endpoint + robust parsing og genforsøg |
| Office.js API-huller (PowerPoint) | Tjek requirement sets ved runtime og skjul funktioner, der ikke understøttes |
| Prompt injection via dokumentindhold | Indhold i afgrænsede tags; agent-tilstand kræver bekræftelse |
