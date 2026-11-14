# MIAI – AI-assistent til Office med dine egne modeller

MIAI er en tilføjelse til Microsoft **Word, Excel, PowerPoint og Outlook**, der bruger **dine egne sprogmodeller**. Den taler med ethvert inference-værktøj med et OpenAI-kompatibelt API, fx Ollama, LM Studio, llama.cpp, vLLM, LocalAI eller Jan. Som standard må den kun kalde modeller på din egen maskine, så dokumenterne forlader den ikke.

> **Status:** under aktiv udvikling (0.1.x). Funktionerne er testet automatisk mod en simuleret udgave af Office. Den manuelle afprøvning i rigtig Office på Windows og Mac er ikke afsluttet ([testlisten](docs/UDVIKLING.md#fase-0--manuel-testliste)), så regn med fejl, og [opret gerne et issue](https://github.com/mor10w/MIAIPlugin/issues), når du finder en.

## Hvad kan den

| | Word | Excel | PowerPoint | Outlook |
|---|:-:|:-:|:-:|:-:|
| Ret stavning og grammatik (i Word som sporede ændringer) | ✅ | | ✅ | ✅ |
| Omskriv i anden tone · parafrasér · forenkl | ✅ | | ✅ | ✅ |
| Oversæt · opsummér | ✅ | ✅ ¹ | ✅ | ✅ |
| Skriv artikel, salgstekst eller disposition | ✅ | | ✅ | ✅ ² |
| Lav diagram ud fra tekst | ✅ | | ✅ | |
| Tekst til slides · design af master og layout | | | ✅ | |
| Forklar formel · lav formel · analyser data | | ✅ | | |
| Svar på mail (som udkast, du selv sender) | | | | ✅ |
| Chat med dokumentet, også med billeder og filer | ✅ | ✅ | ✅ | ✅ |
| Agent: beskriv ændringen, godkend planen, fortryd | ✅ | ✅ | ✅ | |

¹ Excel: oversæt. ² Outlook: salgstekst.

- **Agenten** foreslår en plan, og intet ændres, før du har godkendt den. I Excel kan den bl.a. rydde op i data, sortere, formatere tabeller og lave diagrammer og pivottabeller på tværs af ark.
- **Viden fra dine mapper:** med MIAI App eller bridgen kan chatten svare ud fra dine egne filer og vise kilderne.
- **Egne opgaver:** du kan tilføje dine egne handlinger med egen prompt.

Alle funktioner er beskrevet i [brugervejledningen](docs/BRUG.md).

## Installation

Du skal bruge Microsoft 365 eller Office 2021+ på Windows eller Mac, og et inference-værktøj med en model. Har du ikke et i forvejen, er [Ollama](https://ollama.com) det nemmeste:

```bash
ollama pull qwen3:4b
```

Vælg derefter den vej, der passer til dig:

| Vej | Til hvem | Du skal installere |
|---|---|---|
| **1. MIAI App** (anbefales) | Alle, der bare vil i gang | Én app. Den sætter selv MIAI ind i Office. |
| **2. Hostet add-in** | Dig, der ikke vil installere en app | Kun et manifest, og modellen skal tillade kald fra add-in'ets adresse (CORS). |
| **3. Central udrulning** | Organisationer med Microsoft 365 | Intet hos brugerne; it-afdelingen udruller. |
| **4. Fra kildekoden** | Udviklere | Node.js 22 og dette repo. |

### 1. MIAI App (anbefales)

1. Hent den nyeste udgave under [Releases](https://github.com/mor10w/MIAIPlugin/releases/latest):
   - Mac med Apple Silicon (M1 og nyere): `MIAI-<version>-arm64.dmg`
   - Mac med Intel: `MIAI-<version>-x64.dmg`
   - Windows: `MIAI-Setup-<version>-x64.exe`
2. Installér og start appen.
   - **Mac:** træk MIAI til Programmer og åbn den. Appen er endnu ikke signeret med et Apple-certifikat, så macOS afviser den første gang: åbn Systemindstillinger → Anonymitet & sikkerhed og vælg **Åbn alligevel**.
   - **Windows:** kør installeren. SmartScreen advarer, fordi den ikke er signeret endnu: vælg **Flere oplysninger → Kør alligevel**. Installationen kræver ikke administrator.
3. Sig ja til **Gør det nu**, når appen tilbyder at betro sit certifikat og installere MIAI i Office.
4. Genstart Word, Excel og PowerPoint. Knappen **AI-assistent** ligger på fanen **Hjem**.

MIAI ligger derefter i menulinjen (Mac) eller ved uret (Windows). Outlook, opdateringer og fejlfinding er beskrevet i [brugervejledningen](docs/BRUG.md#a-miai-app-anbefales).

### 2. Hostet add-in

Add-in'et ligger på `https://mor10w.github.io/MIAIPlugin/` og kalder selv din model. Du henter [manifestet](https://mor10w.github.io/MIAIPlugin/manifest.xml), lægger det ind i Office og tillader kald fra `https://mor10w.github.io` i dit inference-værktøj. Se [trinene i brugervejledningen](docs/BRUG.md#b-hostet-add-in-uden-app).

### 3. Central udrulning i en organisation

Organisationen hoster add-in'et selv, låser hvilke modeller der må bruges, og udruller det fra Microsoft 365 Admin Center, også til Outlook og Office på web. Se [Central udrulning](docs/BRUG.md#central-udrulning-i-en-organisation-også-office-på-web).

### 4. Fra kildekoden

```bash
npm install
```

```bash
npm run certs
```

```bash
npm run dev
```

Læg derefter udviklingsmanifestet ind i Office med `npm run sideload:mac` eller `npm run sideload:windows`, og genstart Office. Resten står i [Udvikling og afprøvning](docs/UDVIKLING.md).

## Sådan virker det

```
Office (task pane) ──OpenAI API──────────────────────────────▶ dit inference-værktøj     (direkte)
Office (task pane) ──https://localhost──▶ MIAI Bridge / App ──▶ dit inference-værktøj     (bridge)
```

- **Direkte:** add-in'et er en statisk webside (fx på GitHub Pages), der selv kalder modellen fra Office. Det kræver, at inference-værktøjet tillader CORS fra add-in'ets adresse.
- **Bridge:** en lokal service serverer add-in'et og kalder modellen. Den giver låst politik for organisationer, audit-log, nøgler i systemets nøglering, udtræk af PDF- og Office-filer og viden fra dine mapper. **MIAI App** er bridgen pakket som en app.

Begge tilstande bruger samme motor og samme **politik** for, hvilke hosts der må kaldes (standard: kun `127.0.0.1` og `localhost`). Tilstanden vælges automatisk. Se [arkitekturen](docs/ARKITEKTUR.md) og [ADR-0007](docs/adr/0007-bridge-valgfri-direkte-forbindelse.md).

## Dokumentation

- [Brugervejledning](docs/BRUG.md): installation, indstillinger, alle funktioner og fejlfinding
- [Udvikling og afprøvning](docs/UDVIKLING.md): byg, tests, udgivelse og testlister
- [Testvejledning](docs/TESTVEJLEDNING.md): trin for trin-afprøvning i rigtig Office
- [Arkitektur](docs/ARKITEKTUR.md) og [arkitekturbeslutninger (ADR)](docs/adr/)
- [Eval-sæt og modelvalg](docs/EVAL.md)
- [Brugerflade og designsystem](docs/ui/README.md)
- [Projektplan](docs/PLAN.md)

## Struktur

```
packages/shared   fælles motor: opgaver og prompts, OpenAI-klient, politik, agent
packages/addin    task pane til Office (React, Fluent UI, Office.js) og manifester
packages/bridge   valgfri lokal Node.js-service (Fastify)
packages/app      MIAI App: menulinje-app (Electron), der kører bridgen
packages/evals    eval-sæt til sammenligning af modeller
```

## Bidrag

Fejlrapporter og forslag er velkomne som [issues](https://github.com/mor10w/MIAIPlugin/issues). Vil du ændre i koden, så læs [CONTRIBUTING.md](CONTRIBUTING.md) først. Projektets sprog er dansk.

## Licens

MIAI er udgivet under [MIT-licensen](LICENSE). Ophavsretten tilhører **ManInfo ApS**.

Du må frit bruge, ændre og videregive koden, også i kommercielle produkter og uden aftale med ManInfo ApS. Den eneste betingelse er, at ophavsretsnotitsen og licensteksten følger med i kopier af koden eller væsentlige dele af den. Softwaren leveres, som den er, uden nogen form for garanti.
