# Udvikling og afprøvning

## Forudsætninger
- Node.js 22 LTS
- Microsoft 365 / Office 2021+ desktop (Windows eller Mac)
- Et OpenAI-kompatibelt inference-værktøj, fx [Ollama](https://ollama.com):
  ```bash
  ollama pull qwen3:4b
  ```
  Alternativer: LM Studio (start "Local Server"), `llama-server -m model.gguf --port 8080`, vLLM, LocalAI, Jan m.fl.

## Brug uden at bygge noget (hostet add-in, direkte forbindelse)
Når GitHub Pages er slået til (Settings → Pages → Source: **GitHub Actions**), publiceres add-in'et ved hvert push til `main`.

1. Hent `https://<bruger>.github.io/<repo>/manifest.xml`.
2. Tillad CORS fra `https://<bruger>.github.io` i dit værktøj. For Ollama:
   - Windows: sæt brugermiljøvariablen `OLLAMA_ORIGINS=https://mor10w.github.io` og genstart Ollama.
   - Mac: `launchctl setenv OLLAMA_ORIGINS https://mor10w.github.io` og genstart Ollama.

   LM Studio: slå "Enable CORS" til.
3. Sideload manifestet (se nedenfor) og genstart Office.

## Udvikling – direkte forbindelse (standard)
```bash
npm install
```
```bash
npm run certs
```
Opretter et udviklingscertifikat for `localhost` og gør CA'en betroet (første gang kræves bekræftelse).
```bash
npm run dev
```
Starter Vite på `https://localhost:3000` med hot reload. Ollama tillader `https://localhost:*` som standard, så der er ingen CORS-opsætning.

Sideload `packages/addin/manifest.dev.xml` (én gang):
```bash
npm run sideload:mac
```
```bash
npm run sideload:windows
```

## Udvikling – med bridge (valgfri)
```bash
npm run build
```
```bash
npm start
```
Bridgen kører på `https://localhost:44321` og serverer også task pane. Sideload bridge-manifestet:
```bash
npm run sideload:mac -- packages/addin/manifest.xml
```
På Windows: `npm run sideload:windows -- packages\addin\manifest.xml`.

Watch-tilstand for både add-in og bridge: `npm run dev:bridge`.

## Tests
```bash
npm test
```
- `shared`: opgaver, tekst-hjælpere og motoren i direkte tilstand mod en mock-server.
- `bridge`: streaming, `<think>`-filter, JSON-validering med genforsøg, politik og sikkerhedstjek.

Browsertest (manuel, Playwright/Chromium): add-in'et serveres over HTTPS med et simuleret Office.js og en mock-model, i både direkte og bridge-tilstand. Se ADR-0007 for resultater.

## Manifester
```bash
npm run manifests
```
Regenererer `manifest.xml` (bridge) og `manifest.dev.xml` fra `manifest.template.xml`. Hostet variant:
```bash
node scripts/make-manifest.mjs hosted https://intern.server/miai/ manifest.xml
```

## Konfiguration
**Direkte tilstand**: indstillinger gemmes i add-in'ets `localStorage` på maskinen, og de redigeres under fanen Indstillinger.

**Bridge**:

| | Sti |
|---|---|
| Brugerindstillinger | Windows `%APPDATA%\MIAI\config.json` · Mac `~/Library/Application Support/MIAI/config.json` · Linux `~/.config/miai/config.json` |
| Maskinpolitik (organisation) | Windows `%ProgramData%\MIAI\policy.json` · Mac `/Library/Application Support/MIAI/policy.json` · Linux `/etc/miai/policy.json` |
| Audit-log (hvis slået til) | `audit.log` ved siden af `config.json` |

Miljøvariabler for bridgen: `MIAI_PORT` (44321), `MIAI_CONFIG_DIR`, `MIAI_POLICY_FILE`, `MIAI_TLS_CERT`, `MIAI_TLS_KEY`, `MIAI_ADDIN_DIR`, `MIAI_LOG`.

Eksempel på maskinpolitik, der kun tillader en intern inference-server:
```json
{ "allowedHosts": ["127.0.0.1", "localhost", "llm.firma.local"], "locked": true }
```

## Fase 0 – manuel testliste
Skal afprøves i Office desktop.

**Forbindelse**
- [ ] Direkte, dev (`https://localhost:3000`) → Ollama uden CORS-opsætning, i Windows (WebView2) og Mac (WKWebView)
- [ ] Direkte, hostet (GitHub Pages) → Ollama med `OLLAMA_ORIGINS`: kommer der en Local Network Access-prompt i WebView2?
- [ ] Direkte → LM Studio med og uden "Enable CORS": fejlteksten skal forklare opsætningen
- [ ] Bridge-tilstand detekteres automatisk; badge viser "bridge"
- [ ] `localStorage` bevares mellem genstarter af Word/Excel/PowerPoint (især på Mac)

**Funktioner**
- [ ] Word: ret grammatik → sporede ændringer vises og kan accepteres/afvises
- [ ] Word: omskriv, oversæt, opsummér (uden markering = hele dokumentet), artikel indsættes som formateret tekst, kommentar
- [ ] Excel: forklar formel; generér formel → validering og **Fortryd**; oversæt område → celler; kommentar
- [ ] PowerPoint: omskriv tekst i markeret tekstboks; tekst til slides med standardskabelon **og** en firmaskabelon
- [ ] Chat i alle tre apps; stop-knappen afbryder generering
- [ ] Endpoint uden for politik afvises; låst maskinpolitik (bridge) kan ikke ændres
- [ ] Mørkt Office-tema
- [ ] Reasoning-model (Qwen3): ingen `<think>` i output
