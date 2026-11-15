# Testvejledning: Office, webviews og modelvalg (#1, #9, #2)

Vejledningen beskriver de tre tests, der kræver rigtig Office og rigtig hardware og derfor ikke kan køre i CI:

| Issue | Hvad | Kræver | Tid |
|---|---|---|---|
| #9 | Direkte tilstand i WebView2 (Windows) og WKWebView (Mac) med det hostede add-in | Windows 11 og Mac med Microsoft 365, Ollama | ca. 30 min. |
| #1 | End-to-end i Word, Excel og PowerPoint på begge platforme | Samme som #9 | 2–3 timer pr. platform |
| #2 | Eval-sættet mod rigtige modeller og en anbefaling pr. kategori | En maskine med de relevante modeller (gerne én laptop og én med GPU) | 1–4 timer pr. maskine |

Kør dem i rækkefølgen #9 → #1 → #2. #9 afgør, hvilken tilstand (direkte eller bridge) #1 skal testes i, og #1 viser, hvilke modeller der i praksis føles hurtige nok, før #2 måler dem.

Indhold:
- [Forberedelse](#forberedelse)
- [Fejlfinding og logs](#fejlfinding-og-logs)
- [#9: Direkte tilstand i webviewet](#9-direkte-tilstand-i-webviewet)
- [#1: End-to-end i Office](#1-end-to-end-i-office)
- [#2: Eval og modelvalg](#2-eval-og-modelvalg)

---

> **Kort udgave:** vil du blot vide, om MIAI kan det, demovideoen viser, så brug [tjeklisten i `docs/demo/`](demo/TJEKLISTE.md). Den gentager videoens fire forløb i rigtig Word og Excel med færdige testfiler.

## Forberedelse

### På begge maskiner
1. **Microsoft 365** (Office-version 2408 eller nyere). Notér versionen: *Filer → Konto → Om Word*.
2. **Ollama** fra [ollama.com](https://ollama.com). Hent to modeller, en lille og en standard:
   ```bash
   ollama pull qwen3:4b
   ```
   ```bash
   ollama pull gemma3:12b
   ```
   Til F15 (videnbase) også en embedding-model:
   ```bash
   ollama pull embeddinggemma:300m
   ```
3. **Node.js 22.13 eller nyere** og repoet, kun til bridge-test og eval:
   ```bash
   git clone https://github.com/mor10w/MIAIPlugin.git
   ```
   ```bash
   cd MIAIPlugin && npm install && npm run build
   ```
4. **Testdokumenter**, se [Testdata](#testdata) nedenfor. Lav dem én gang og genbrug dem på begge platforme.

### Tilstande, der skal testes
| Tilstand | Manifest | Hvornår |
|---|---|---|
| **A. Direkte, hostet** | `https://mor10w.github.io/MIAIPlugin/manifest.xml` | #9 og som standard i #1 |
| **B. Bridge** | Personligt manifest fra `npm start` (stien vises ved start) | Fallback, hvis #9 viser, at A ikke virker |
| **C. MIAI App** | Installeres af appen selv | Kort røgtest i #1. Tjeklisten i [UDVIKLING.md](UDVIKLING.md#fase-0--manuel-testliste) |

Hav kun ét MIAI-manifest sideloadet ad gangen, ellers vises to knapper med samme navn. Fjern et manifest igen:
- Windows: `Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name <Id>`.
- Mac: slet `miai-*.xml` i `~/Library/Containers/com.microsoft.{Word,Excel,Powerpoint}/Data/Documents/wef/`.

Genstart derefter Office.

---

## Fejlfinding og logs

**Udviklerværktøjer i task pane**
- **Windows (WebView2):** højreklik i task pane → *Undersøg* (Inspect). Kommer menuen ikke frem, så åbn *Microsoft Edge DevTools Preview* fra Microsoft Store og vælg add-in'ets mål.
- **Mac (WKWebView):** slå udviklermenuen til én gang pr. app og genstart appen:
  ```bash
  defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true
  ```
  ```bash
  defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
  ```
  ```bash
  defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true
  ```
  Højreklik derefter i task pane → *Inspicér element*. Safari's webinspektør åbner.

**Miljøet** (indsæt i konsollen i udviklerværktøjerne, og kopiér resultatet til rapporten):
```js
const d = Office.context.diagnostics;
const sets = { WordApi: ["1.3", "1.4", "1.6"], ExcelApi: ["1.9", "1.10"], PowerPointApi: ["1.2", "1.4", "1.8"], Mailbox: ["1.5"] };
[`${d.host} ${d.platform} ${d.version}`, navigator.userAgent,
 ...Object.entries(sets).map(([s, v]) => `${s}: ${v.map((x) => `${x}=${Office.context.requirements.isSetSupported(s, x) ? "ja" : "nej"}`).join(" ")}`)].join("\n");
```
Det er de requirement sets, MIAI slår funktioner til og fra efter (PLAN.md §10). Kun værtens egne sets giver mening; de andre svarer "nej".

**Logs**
- Bridge: konsollen, hvor `npm start` kører. `MIAI_LOG=debug npm start` viser hver forespørgsel.
- MIAI App: `MIAI.log` i konfigurationsmappen (Windows `%APPDATA%\MIAI\`, Mac `~/Library/Application Support/MIAI/`).
- Ollama: Windows `%LOCALAPPDATA%\Ollama\server.log`, Mac `~/.ollama/logs/server.log`.

**Office cacher add-ins.** Ser en ændring ikke ud til at slå igennem, så ryd cachen og genstart Office:
- Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
- Mac: `~/Library/Containers/com.microsoft.<App>/Data/Library/Caches/`

---

## #9: Direkte tilstand i webviewet

**Spørgsmålet:** Kan add-in'et, når det hentes fra `https://mor10w.github.io`, kalde `http://127.0.0.1:11434` direkte? Og spørger webviewet om lov først?
- Chromium (og dermed WebView2) er ved at indføre *Local Network Access*: et offentligt site skal have brugerens tilladelse til at kalde maskinen selv eller det lokale net.
- WebKit (Mac) har sine egne regler for HTTP fra en HTTPS-side.

### Opsætning (begge platforme)
1. Tillad add-in'ets adresse i Ollama:
   - Windows: *Systemegenskaber → Miljøvariabler → Brugervariabler → Ny*, navn `OLLAMA_ORIGINS`, værdi `https://mor10w.github.io`. Afslut Ollama i proceslinjen og start den igen.
   - Mac:
     ```bash
     launchctl setenv OLLAMA_ORIGINS https://mor10w.github.io
     ```
     Afslut Ollama i menulinjen og start den igen.
2. Kontrollér fra en terminal, at Ollama svarer med CORS-headeren:
   ```bash
   curl -si -H "Origin: https://mor10w.github.io" http://127.0.0.1:11434/v1/models | grep -i access-control
   ```
   Forventet: `Access-Control-Allow-Origin: https://mor10w.github.io`. Uden den linje er fejlen i Ollamas opsætning, ikke i webviewet.
3. Hent og sideload det hostede manifest:
   - Windows (PowerShell, i repoet):
     ```powershell
     Invoke-WebRequest https://mor10w.github.io/MIAIPlugin/manifest.xml -OutFile $env:USERPROFILE\miai-hosted.xml
     ```
     ```powershell
     npm run sideload:windows -- $env:USERPROFILE\miai-hosted.xml
     ```
   - Mac:
     ```bash
     curl -so ~/Downloads/miai-hosted.xml https://mor10w.github.io/MIAIPlugin/manifest.xml
     ```
     ```bash
     npm run sideload:mac -- ~/Downloads/miai-hosted.xml
     ```
   Uden repoet kan scriptet erstattes af `reg add` (Windows) eller en kopi til wef-mapperne (Mac), se [Tilstande](#tilstande-der-skal-testes).
4. Luk alle Office-apps helt, og åbn Word.

### Test (Word på Windows, derefter Word på Mac)
| # | Trin | Notér |
|---|---|---|
| 9.1 | *Hjem → AI-assistent*. | Åbner task pane? Badge'et viser "direkte"? |
| 9.2 | Første kald: Indstillinger → *Test forbindelse* (eller vælg en opgave). | **Vises der en dialog om lokalt netværk eller enheder på netværket?** Ordlyd og skærmbillede. |
| 9.3 | Hvis ja: vælg *Tillad*. | Virker kaldet bagefter? Huskes valget efter genstart af Word? Gælder det også Excel og PowerPoint? |
| 9.4 | Hvis ja: gentag med en ren cache, og vælg *Bloker*. | Hvilken fejltekst viser MIAI? Kan man fortryde blokeringen, og hvor? |
| 9.5 | Hvis nej: kører kaldet uden videre? | Skærmbillede af en fuldført opgave. |
| 9.6 | Kør i konsollen i udviklerværktøjerne: `await fetch("http://127.0.0.1:11434/v1/models").then(r => r.status)` | 200, eller fejlbeskeden. Isolerer webviewet fra MIAI's egen kode. |
| 9.7 | Prøv `http://localhost:11434/v1` som endpoint i Indstillinger. | Samme adfærd som `127.0.0.1`? |
| 9.8 | Fjern `OLLAMA_ORIGINS`, og genstart Ollama. | MIAI's fejltekst skal forklare CORS-opsætningen. |
| 9.9 | LM Studio med og uden *Enable CORS* (valgfrit). | Som 9.8. |
| 9.10 | Mac: afslut Word helt og åbn igen. | Er indstillingerne (`localStorage`) bevaret? |

På Mac er spørgsmålet i 9.2 i stedet: virker kaldet overhovedet (WKWebView og blandet indhold), og beder macOS om tilladelse til *lokalt netværk* (Systemindstillinger → Anonymitet og sikkerhed → Lokalt netværk)?

### Hvis direkte tilstand ikke virker
Prøv fallbacks i denne rækkefølge, og notér hvilke der virker:
1. **Bridge** (`npm start`, sideload det personlige manifest). Alt går via `https://localhost:44321`, så der er ingen kald fra offentlig til lokal adresse.
2. **MIAI App**, som er bridgen pakket som app.
3. **Dev-varianten** (`npm run dev`, `manifest.dev.xml`, task pane fra `https://localhost:3000`) som kontrol: virker den, men ikke den hostede, er problemet origin'en og ikke CORS.

### Resultat (indsæt som kommentar i #9)
```markdown
**Platform:** Windows 11 23H2 / macOS 15.x · **Office:** Version 25xx (Build …) · **WebView2/Safari:** <fra navigator.userAgent>
**Requirement sets:** <output fra konsollen>

| Trin | Resultat |
|---|---|
| 9.2 Local Network Access-prompt | ja/nej – <ordlyd>, <skærmbillede> |
| 9.3 Tillad huskes | ja/nej, gælder alle apps ja/nej |
| 9.4 Bloker → fejltekst | <tekst> |
| 9.6 fetch fra konsollen | 200 / <fejl> |
| 9.7 localhost vs. 127.0.0.1 | ens / forskel: … |
| 9.8 Uden OLLAMA_ORIGINS | fejltekst forståelig ja/nej |
| 9.10 localStorage bevaret | ja/nej |
| Fallback, der virker | bridge / app / dev |

**Konsekvens:** <fx "direkte tilstand virker efter én prompt pr. app"; "kræver bridge på Mac">
```
Når resultatet er kendt, skal BRUG.md (afsnittet om direkte forbindelse) beskrive prompten og hvad man gør ved den. ADR-0007 og PLAN.md §10 (risici) skal opdateres med konklusionen.

---

## #1: End-to-end i Office

Testmatrixen er **app × platform × funktion**. Udfør den i tilstand A, eller i B, hvis #9 viste, at A ikke virker. Brug `qwen3:4b` til det meste, så det går hurtigt, og `gemma3:12b` til artikel, slides og agent.

Start hver app med at køre konsol-snippet'et fra [Fejlfinding](#fejlfinding-og-logs) og gem resultatet (acceptkriteriet om requirement sets).

### Testdata
**Word**, `test-word.docx`: en titel (typografi *Titel*), to overskrifter (*Overskrift 1*), et afsnit med **fed tekst** og et link, en punktliste med tre punkter, en nummereret liste og en tabel 3×3. Brug denne tekst med vilje indsatte fejl:

> Vores firma har i år øget omsætningen med 12 procent, hvilket er det bedste resultat siden 2019. Vi har ansat fem nye medarbejder og åbnet et kontor i Aarhus. Det næste år vil vi fokusere på at forbedre kundeoplevelsen og at udvide til Sverrig.

**Excel**, `test-excel.xlsx`: indsæt nedenstående i A1, og gør det til en tabel (*Indsæt → Tabel*) med kr.-format i C og procent i D. Gør overskriften fed med en fyldfarve.

```
Region	Måned	Omsætning	Margin
Nord	Januar	125000	0,18
Syd	Januar	98000	0,22
Nord	Februar	131000	0,17
Syd	Februar	104500	0,21
Øst	Januar	76000	0,25
```

**PowerPoint**, `test-ppt.pptx`: to præsentationer.
1. Standardtemaet med tre slides: titel, punktliste med et fedt ord og farvet tekst, og et slide med billede og noter.
2. En firmaskabelon (en `.potx` fra en kunde eller et indbygget tema som *Ion* eller *Retrospekt*).

### Matrix
Markér hver celle **OK**, **Afvigelse** (med issue-nummer) eller **–** (ikke relevant). Funktionsnumrene følger PLAN.md §3.

| Funktion | Word Win | Word Mac | Excel Win | Excel Mac | PPT Win | PPT Mac |
|---|---|---|---|---|---|---|
| F1 Ret grammatik (sporede ændringer) | | | – | – | | |
| F2 Omskriv, parafrasér, forenkl, tone | | | – | – | | |
| F3 Oversæt markering | | | | | | |
| F4 Chat om dokumentet, stop-knap | | | | | | |
| F5 Artikel | | | – | – | – | – |
| F6 Salgstekst | | | – | – | | |
| F7 Forklar formel | – | – | | | – | – |
| F8 Generér formel, validering og Fortryd | – | – | | | – | – |
| F9 Disposition | | | – | – | | |
| F10 Tekst → slides (standard og firmaskabelon) | – | – | – | – | | |
| F11 Opsummér (med og uden markering) | | | – | – | | |
| F12 Dataanalyse | – | – | | | – | – |
| F13 Agent og Fortryd | | | | | | |
| F14 Egen opgave fra fil | | | | | | |
| F15 Videnbase (kun bridge eller app) | | | | | | |
| Design: master og layout | – | – | – | – | | |
| Mørkt Office-tema | | | | | | |
| Reasoning-model: ingen `<think>` i output | | | | | | |

### Fremgangsmåde pr. funktion
Det skal især tjekkes:

| Funktion | Gør | Forventet |
|---|---|---|
| F1 | Markér afsnittet med fejl → *Ret stavning og grammatik*. | Word: sporede ændringer ("medarbejdere", "Sverige"), kan accepteres og afvises enkeltvis. Rigtig tekst rettes ikke. PowerPoint: markeringen erstattes. |
| F2 | Markér et afsnit med fed tekst → *Omskriv i anden tone* (formel). | Teksten skifter; formatering, der ikke berøres, bevares. Fortryd-knappen gendanner. |
| F3 | Word: markér to afsnit → engelsk. Excel: markér A1:B6 → engelsk. | Word: oversættelse på stedet. Excel: hver celle oversat, tal og formater urørte. |
| F4 | Spørg "Hvor har vi åbnet kontor?" og derefter "Hvem er direktør?". | Første svar fra dokumentet; andet svar "står ikke i dokumentet". Stop afbryder streaming med det samme. |
| F5 | *Skriv artikel ud fra idé*: "Hybridarbejde i små virksomheder". | Indsættes som formateret tekst med overskrifter i Words typografier, ikke som Markdown-tegn. |
| F7/F8 | Excel: *Lav formel ud fra beskrivelse* i F2: "Summen af omsætning for Nord". | Formlen bruger `SUMIFS`, og cellen viser 256000. *Fortryd* fjerner formlen og gendanner cellens format. Ctrl+Z virker ikke på add-in-ændringer (PLAN.md §7), så Fortryd i MIAI skal virke. |
| F10 | PowerPoint: indsæt to afsnit tekst → *Tekst til slides*. | Slides oprettes efter det markerede slide og med layouts fra temaet. Gentag med firmaskabelonen. |
| F11 | Uden markering → *Opsummér*. | Bruger hele dokumentet (der står "hele dokumentet" i opgaven). Med markering kun markeringen. |
| F12 | Markér tabellen → *Analyser data*. | Nøgletal stemmer (Nord 256.000), forslaget til pivot/diagram giver mening. |
| F13 | Brug tjeklisterne *Agent i Word/Excel/PowerPoint – formatering* i [UDVIKLING.md](UDVIKLING.md#fase-0--manuel-testliste). | Plan vises før udførelse, formatering bevares, Fortryd gendanner. |
| F14 | Fanen *Opgaver*: opret en opgave, og importér en JSON-fil (BRUG.md). | Opgaven vises under sit navn og kan køres; eksport giver samme fil tilbage. |
| F15 | Bridge eller app: tilføj en mappe med et par .docx/.pdf → spørg om indholdet. | Svaret citerer kilden. |
| Design | Tjeklisten *Design: master og layout* i UDVIKLING.md. | |

Kør også de øvrige punkter under *Forbindelse* og *Funktioner* i [UDVIKLING.md](UDVIKLING.md#fase-0--manuel-testliste) og, hvis du har Outlook, afsnittet *Outlook og central udrulning (#7)*.

### Hvis noget fejler
1. Gentag én gang for at udelukke modellens tilfældighed. Fejler det igen med `gemma3:12b`, er det sandsynligvis en fejl i MIAI.
2. Gem konsollens fejl (udviklerværktøjerne), bridge-loggen og et skærmbillede.
3. Opret et issue med titlen `[#1] <App> <platform>: <kort beskrivelse>`, og angiv Office-version, tilstand (A/B/C), model og trin.

### Resultat (indsæt som kommentar i #1)
```markdown
**Platform:** … · **Office:** … · **Tilstand:** A/B/C · **Modeller:** qwen3:4b, gemma3:12b

**Requirement sets**
- Word: …
- Excel: …
- PowerPoint: …

<matrixen udfyldt>

**Afvigelser:** #xx, #yy
```

---

## #2: Eval og modelvalg

Eval-sættet og harnesset er færdige (se [EVAL.md](EVAL.md)). Tilbage er kørslen på rigtig hardware og anbefalingen.

### Hardware
Kør helst på to profiler, så anbefalingen dækker både laptop og GPU:

| Profil | Eksempel | Modeller |
|---|---|---|
| Laptop uden GPU | 16 GB RAM, nyere Intel/AMD eller Apple M1/M2 | `qwen3:4b`, `gemma3:4b`, `phi4-mini` |
| Arbejdsstation eller GPU | 32 GB+ RAM eller GPU med 16–24 GB | `gemma3:12b`, `qwen3:14b`, `mistral-small3.2`, `qwen2.5-coder:14b` |
| Dommer | Den største, der kan være der | `qwen3:32b` eller `gemma3:27b` |

Notér CPU/GPU, RAM, OS, Ollama-version (`ollama -v`) og kvantisering (standard `Q4_K_M`, se `ollama show <model>`).

### Forberedelse
1. Hent modellerne, fx:
   ```bash
   for m in qwen3:4b gemma3:4b phi4-mini; do ollama pull $m; done
   ```
2. Hæv Ollamas kontekstvindue, så lange cases ikke afkortes. Sæt det som miljøvariabel som `OLLAMA_ORIGINS` ovenfor, og genstart Ollama:
   ```bash
   launchctl setenv OLLAMA_CONTEXT_LENGTH 8192
   ```
   (Windows: brugervariablen `OLLAMA_CONTEXT_LENGTH=8192`.)
3. Kontrollér opsætningen med en kort kørsel:
   ```bash
   npm run eval -- --base-url http://127.0.0.1:11434/v1 --model qwen3:4b --filter formula
   ```
   Brug `--list` for at se cases og filtre.

### Kørsler
1. **Laptop, hele sættet, tre gentagelser** (stabilitet):
   ```bash
   npm run eval -- --base-url http://127.0.0.1:11434/v1 --model qwen3:4b --model gemma3:4b --model phi4-mini --repeat 3
   ```
2. **Standardmodeller med dommer:**
   ```bash
   npm run eval -- --base-url http://127.0.0.1:11434/v1 --model gemma3:12b --model qwen3:14b --model mistral-small3.2 --judge-model qwen3:32b
   ```
3. **Formler og struktureret output**, med kodemodellen:
   ```bash
   npm run eval -- --base-url http://127.0.0.1:11434/v1 --model qwen3:14b --model qwen2.5-coder:14b --filter formula --structured json_schema --structured json_object --repeat 3
   ```
4. **Gentagelsesstraf** på den model, der vinder dansk skrivning (tjek, om dansken holder ved 1,5):
   ```bash
   npm run eval -- --base-url http://127.0.0.1:11434/v1 --model gemma3:12b --presence 0 --presence 1 --presence 1.5
   ```
5. Valgfrit: samme modeller i MLX (oMLX eller LM Studio på Mac) for at sammenligne hastighed.

Lad maskinen være i fred under kørslen; andre programmer forstyrrer tokens/s.

### Fra rapport til anbefaling
1. Kopiér rapporterne fra `packages/evals/results/` til `docs/eval/` med navnet `ÅÅÅÅ-MM-DD-<profil>.md` (fx `2026-11-10-laptop-m2-16gb.md`). JSON-rådata committes ikke.
2. Vælg pr. kategori (rapportens *bedste model pr. kategori*) efter denne rækkefølge:
   1. Bestået ≥ 80 %.
   2. Struktureret: JSON 1. forsøg ≥ 90 %.
   3. Gentagelser ≤ 2 %.
   4. Højeste dommerscore.
   5. Ved lige stand: hurtigste median.
3. Læs 5–10 danske svar fra vinderen i JSON-rådata. Automatiske tjek fanger ikke klodset dansk.
4. Skriv anbefalingen i EVAL.md under *Anbefalinger*: en tabel med kategori × profil → model, og en linje med hardware, Ollama-version og dato. Opdatér PLAN.md §5 tilsvarende. Opgavekatalogets anbefalinger i UI'et (#13) skal følge samme tabel.

### Resultat (indsæt som kommentar i #2)
```markdown
**Hardware:** … · **Ollama:** 0.x · **Kvantisering:** Q4_K_M · **Rapporter:** docs/eval/…

| Kategori | Laptop | Arbejdsstation/GPU |
|---|---|---|
| Skrivning (dansk) | | |
| Oversættelse | | |
| Resumé og analyse | | |
| Struktureret (JSON) | | |
| Formler | | |
| Chat | | |

**Bemærkninger:** <fx "phi4-mini skriver dårligt dansk", "presence 1,5 giver engelske ord">
```
