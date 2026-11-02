# MIAIPlugin – projektplan (udkast v0.2)

> Status: **Udkast** · Senest opdateret: 2026-11-02 · Ejer: M. Wiingaard

## 1. Vision

En AI-assistent til Microsoft Word, Excel og PowerPoint (senere Outlook). Den
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
| Installation | Installer pr. bruger | Pakket til Intune/SCCM, Jamf |

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

Detaljer i [ARKITEKTUR.md](ARKITEKTUR.md).

```
┌──────────── Office desktop (Word / Excel / PowerPoint) ─────────────┐
│ Task pane (React + Fluent UI)  ←  serveres af bridgen (same origin) │
│ Host-adaptere: Word · Excel · PowerPoint (Office.js)                │
└───────────────┬─────────────────────────────────────────────────────┘
                │ https://localhost:44321/api/*  (SSE-streaming)
┌───────────────▼───────────── MIAI Bridge (Node.js) ─────────────────┐
│ Opgavekatalog/prompts · politik (allowedHosts) · indstillinger      │
│ <think>-filter · JSON-output m. validering · audit-log              │
└───────────────┬─────────────────────────────────────────────────────┘
                │ OpenAI API: /v1/models, /v1/chat/completions
     Ollama · LM Studio · llama.cpp · vLLM · LocalAI · Jan · …
```

## 5. Modelstrategi

Modelvalg er konfiguration. Bridgen lister modeller fra endpointen
(`/v1/models`); brugeren vælger en standardmodel og evt. en model pr. opgave.

| Brug | Forslag (evalueres i fase 0) | Ca. krav (Q4) |
|------|------------------------------|-------------|
| Standard, dansk | Gemma 3 12B, Qwen3 14B, Mistral Small 3.x | 10–16 GB |
| Laptop uden GPU | Qwen3 4B, Gemma 3 4B, Phi-4-mini | 4–6 GB |
| Formler, struktureret output | Qwen3, Qwen2.5-Coder | – |

Struktureret output (grammatik-rettelser, slides, formler) bruger
`response_format: json_schema`, når endpointen understøtter det. Ellers bruges
`json_object` eller prompt + parsing + ét genforsøg. Dette sættes pr. endpoint
(`structuredOutput`).

Reasoning-modeller: `<think>…</think>` og `reasoning_content` filtreres fra i bridgen.

## 6. Faseplan

### Fase 0 – Spike (påbegyndt)
- [x] Monorepo-skelet: `shared`, `bridge`, `addin`.
- [x] Bridge: HTTPS, statiske add-in-filer, `/api/health`, `/api/models`, `/api/tasks/:id` (SSE), `/api/settings`, politik-check.
- [x] Add-in: task pane, host-adaptere, hurtighandlinger, chat, indstillinger.
- [ ] Verificér end-to-end i Word, Excel og PowerPoint desktop på Windows og Mac (manuel test).
- [ ] Eval-sæt (30–50 opgaver, dansk/engelsk) og valg af anbefalede modeller.

### Fase 1 – MVP
- Hærdning af F1–F11, fejlhåndtering, afbryd-knap, tokengrænser og chunking af store dokumenter.
- Installer: Windows (MSI/winget) og macOS (pkg). Installerer bridge som bruger-service, opretter lokalt CA/cert og sideloader manifest.
- Bakkeikon/statusside for bridgen.

### Fase 2
- F6, F12, F14: opgaver som brugerredigerbare filer, deling via mappe.
- Politikfil og audit-log færdiggjort (forberedelse til enterprise).

### Fase 3
- F13: agent-tilstand (tool calling, whitelisted Office-værktøjer, plan → bekræft → udfør).
- Outlook. Office på web via central udrulning plus on-prem endpoint.

### Fase 4
- F15: lokal RAG (SQLite + sqlite-vec, embeddings via `/v1/embeddings` på samme endpoint).

## 7. Ikke-funktionelle krav

- **Privatliv**: bridgen kalder kun hosts, som politikken tillader; ingen telemetri.
- **Sikkerhed**: lytter kun på loopback. Kræver korrekt `Host`-header (DNS rebinding) og eget `Origin`, så fremmede websider ikke kan bruge bridgen. Dokumentindhold markeres som data i prompts.
- **Ydelse**: streaming altid for tekstsvar; første token < 2 s på anbefalet hardware.
- **Fortryd**: Word bruger track changes eller Ctrl+Z. Excel gemmer de tidligere formler og har en "Fortryd"-knap, fordi add-in-ændringer ikke altid kan fortrydes med Ctrl+Z.

## 8. Teknologistak

| Lag | Valg |
|-----|------|
| Add-in | TypeScript, React 19, Fluent UI v9, Vite, Office.js, XML-manifest |
| Bridge | Node.js 22 LTS, TypeScript, Fastify 5, Zod (ADR-0004) |
| Test | Vitest (bridge og shared, mod mock OpenAI-server); senere Playwright og eval-harness |
| CI | GitHub Actions |

## 9. Beslutninger og åbne spørgsmål

Besluttet:
- Kun OpenAI-kompatibel API (ADR-0003).
- Desktop først; privat og SMB først (ADR-0005).
- Bridge i Node.js/TypeScript (ADR-0004).
- Politik/allowedHosts fra start (ADR-0006).

Åbne:
1. Licens: open source (MIT/Apache-2.0) eller kommerciel/dual?
2. Endeligt produktnavn.
3. Kodesignering af installer: certifikat og udgiver.

## 10. Risici

| Risiko | Afbødning |
|--------|-----------|
| Office-webview stoler ikke på lokalt cert | Installer tilføjer CA til brugerens trust store; i udvikling bruges `office-addin-dev-certs` |
| Lokale modeller svagere på dansk | Eval-sæt; model pr. opgave; enterprise kan pege på større on-prem model |
| Svag hardware | Små modeller, anbefaling ud fra RAM |
| Endpoints understøtter struktureret output forskelligt | `structuredOutput`-indstilling pr. endpoint + robust parsing og genforsøg |
| Office.js API-huller (PowerPoint) | Tjek requirement sets ved runtime og skjul funktioner, der ikke understøttes |
| Prompt injection via dokumentindhold | Indhold i afgrænsede tags; agent-tilstand kræver bekræftelse |
