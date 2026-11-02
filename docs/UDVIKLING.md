# Udvikling og afprøvning

## Forudsætninger
- Node.js 22 LTS
- Microsoft 365 / Office 2021+ desktop (Windows eller Mac)
- Et OpenAI-kompatibelt inference-værktøj, fx [Ollama](https://ollama.com):
  ```bash
  ollama pull qwen3:4b
  ```
  Alternativer: LM Studio (start "Local Server"), `llama-server -m model.gguf --port 8080`, vLLM, LocalAI, Jan m.fl.

## Kom i gang

```bash
npm install
```

```bash
npm run certs
```
Opretter et lokalt udviklingscertifikat for `localhost` (`~/.office-addin-dev-certs`) og gør CA'en betroet. Kræver bekræftelse første gang.

```bash
npm run build
```

```bash
npm start
```
Bridgen kører nu på `https://localhost:44321` og serverer også task pane.

Sideload manifestet (én gang):

```bash
npm run sideload:mac
```
På Windows:
```bash
npm run sideload:windows
```

Genstart Word, Excel eller PowerPoint. Knappen **AI-assistent** findes under fanen **Hjem**.

### Udviklingstilstand
```bash
npm run dev
```
Bygger add-in'et ved hver ændring og genstarter bridgen. Genindlæs task pane (højreklik → Genindlæs, eller luk/åbn).

## Tests
```bash
npm test
```
Bridgen testes mod en mock OpenAI-server: streaming, `<think>`-filter, JSON-validering med genforsøg, politik, sikkerhedstjek.

## Konfiguration
| | Sti |
|---|---|
| Brugerindstillinger | Windows `%APPDATA%\MIAI\config.json` · Mac `~/Library/Application Support/MIAI/config.json` · Linux `~/.config/miai/config.json` |
| Maskinpolitik (organisation) | Windows `%ProgramData%\MIAI\policy.json` · Mac `/Library/Application Support/MIAI/policy.json` · Linux `/etc/miai/policy.json` |
| Audit-log (hvis slået til) | `audit.log` ved siden af `config.json` |

Miljøvariabler: `MIAI_PORT` (44321), `MIAI_CONFIG_DIR`, `MIAI_POLICY_FILE`, `MIAI_TLS_CERT`, `MIAI_TLS_KEY`, `MIAI_ADDIN_DIR`, `MIAI_LOG`.

Eksempel på maskinpolitik, der kun tillader en intern inference-server:
```json
{ "allowedHosts": ["127.0.0.1", "localhost", "llm.firma.local"], "locked": true }
```

## Fase 0 – manuel testliste
Skal afprøves i Office desktop; kan ikke testes i CI.

- [ ] Word (Win/Mac): ret grammatik → sporede ændringer vises og kan accepteres/afvises
- [ ] Word: omskriv, oversæt, opsummér (uden markering = hele dokumentet), artikel indsættes som formateret tekst
- [ ] Word: kommentar indsættes på markering
- [ ] Excel: forklar formel; generér formel → validering og **Fortryd**; oversæt område (TSV) → celler
- [ ] Excel: kommentar på celle
- [ ] PowerPoint: omskriv tekst i markeret tekstboks; tekst til slides med standardskabelon **og** en firmaskabelon (pladsholder-heuristik)
- [ ] Chat i alle tre apps, stop-knap afbryder generering
- [ ] Indstillinger: tilføj LM Studio/llama.cpp, test forbindelse, skift model
- [ ] Endpoint uden for politik afvises; låst maskinpolitik kan ikke ændres
- [ ] Mørkt Office-tema
- [ ] Reasoning-model (Qwen3) – ingen `<think>` i output
