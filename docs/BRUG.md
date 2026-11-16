# Brugervejledning

Sådan installerer og bruger du MIAI – AI-assistenten til Word, Excel, PowerPoint og Outlook med dine egne modeller.

> **Status:** under aktiv udvikling (0.1.x). Funktionerne er testet automatisk mod en simuleret udgave af Office; den manuelle afprøvning i rigtig Office er ikke afsluttet. Se [testlisten](UDVIKLING.md#fase-0--manuel-testliste) for, hvad der mangler.

## Hvad du skal bruge

- **Microsoft 365 eller Office 2021+ til skrivebordet** (Windows eller Mac). Nogle funktioner kræver nyere API'er:
  - *Sporede ændringer* og *kommentarer i Word*: WordApi 1.4
  - *Opret slides* og *opsummér hele præsentationen*: PowerPointApi 1.4
- **Et OpenAI-kompatibelt inference-værktøj** med en model, fx:
  - [Ollama](https://ollama.com): `ollama pull qwen3:4b`
  - LM Studio (start "Local Server")
  - `llama-server -m model.gguf --port 8080`, vLLM, LocalAI, Jan m.fl.
- **Node.js 22 LTS** – kun hvis du bygger fra kildekoden (vej C og D).

Dokumentet forlader ikke din maskine: standardpolitikken tillader kun endpoints på loopback (`127.0.0.1`, `localhost`).

## Installation – vælg den vej, der passer til dig

| Vej | Til hvem | Du skal installere |
|---|---|---|
| **A. MIAI App** (anbefales) | Alle, der bare vil i gang | Én app. Den sætter selv MIAI ind i Office. |
| **B. Hostet add-in** | Dig, der ikke vil installere en app | Kun et manifest. Modellen skal tillade kald fra add-in'ets adresse (CORS). |
| **C. Udvikler-tilstand** | Udviklere | Node.js og dette repo. |
| **D. MIAI Bridge fra kildekoden** | Udviklere og servere uden skrivebord | Node.js og dette repo. |

Organisationer, der vil udrulle MIAI til mange brugere, bruger [Central udrulning](#central-udrulning-i-en-organisation-også-office-på-web).

### A. MIAI App (anbefales)

MIAI App er en lille app i menulinjen, der kører MIAI lokalt og selv installerer tilføjelsen i Word, Excel og PowerPoint. Du behøver hverken Node.js eller en webserver.

**1. Hent appen** under [Releases](https://github.com/mor10w/MIAIPlugin/releases/latest):

| Maskine | Fil |
|---|---|
| Mac med Apple Silicon (M1 og nyere) | `MIAI-<version>-arm64.dmg` |
| Mac med Intel | `MIAI-<version>-x64.dmg` |
| Windows | `MIAI-Setup-<version>-x64.exe` |

Er du i tvivl på Mac: Apple-menuen → *Om denne Mac*. Står der "Chip: Apple M…", er det Apple Silicon.

**2. Installér og start**

- **Mac:** åbn `.dmg`-filen, træk **MIAI** til Programmer og start appen. Appen er endnu ikke signeret med et Apple-certifikat, så macOS afviser den første gang. Åbn Systemindstillinger → Anonymitet & sikkerhed, rul ned og vælg **Åbn alligevel**. (På macOS 14 og ældre kan du i stedet højreklikke på appen og vælge *Åbn*.)
- **Windows:** kør installeren. Installationen sker pr. bruger og kræver ikke administrator. Installeren er ikke signeret endnu, så SmartScreen advarer: vælg **Flere oplysninger → Kør alligevel**.

**3. Sæt MIAI ind i Office**

1. Ved første start laver appen et TLS-certifikat for `localhost` og tilbyder at betro det og installere MIAI i Office i ét trin: vælg **Gør det nu**. På Windows spørger systemet, om certifikatet må installeres; vælg **Ja**.
2. Genstart Word, Excel og PowerPoint. Knappen **AI-assistent** ligger på fanen **Hjem** (og under Indsæt → Tilføjelser).
3. Vælg endpoint og model under fanen **Indstillinger** i MIAI; se [Første gang](#første-gang).

Begge trin findes også i appens menu (**Betro certifikat…**, **Installer i Office…**). Dialogen vises kun igen, hvis noget mangler, fx når certifikatet fornyes (ca. hvert andet år). Outlook installeres for sig; se [Outlook](#outlook).

**Menuen** ligger i menulinjen på Mac og ved uret på Windows:

- **Start automatisk ved login.** På Mac skal du måske godkende det under Systemindstillinger → Generelt → Login-emner, og appen skal ligge i Programmer.
- **Vis Outlook-manifest…** åbner mappen med den fil, Outlook skal have.

**Opdateringer**

- **Windows:** nye udgaver hentes automatisk i baggrunden og installeres, næste gang appen startes, eller fra menuen *Genstart og opdater til …*.
- **Mac:** automatisk opdatering kræver en signeret app ([ADR-0016](adr/0016-distribution-signering-opdatering.md)) og virker derfor ikke endnu. Hent den nye `.dmg` under [Releases](https://github.com/mor10w/MIAIPlugin/releases/latest), afslut MIAI, træk den nye udgave til Programmer, og godkend den igen under Anonymitet & sikkerhed. Spørger macOS bagefter om adgang til nøgleringen, så vælg *Tillad altid*.
- **Udviklerudgave (Mac med Apple Silicon):** ud over udgivelserne bygges en udviklerudgave, hver gang koden ændres. Er der en nyere end din, viser menuen *Hent udviklerudgave …*, som henter `.dmg`-filen i browseren; den installeres på samme måde. Den er til afprøvning og kan indeholde fejl.

**Hvis noget driller**

- *Der sker ingenting, når du starter MIAI:* MIAI har intet vindue; den ligger i menulinjen (Mac) eller ved uret (Windows). Kører den allerede, siger den det nu, og er det en anden udgave, kan du skifte til den. Udgaver før 0.1.9 lukker den nye uden et ord: afslut den gamle fra menuen (eller `pkill -x MIAI` i Terminal på Mac), og start igen.
- *Ikonet i Dock bliver ved med at hoppe (Mac):* luk en kørende MIAI som ovenfor. Hjælper det ikke, så se under Systemindstillinger → Anonymitet & sikkerhed, om MIAI er blokeret, og vælg **Åbn alligevel**.
- *"Adgang nægtet" ved installation i Office (Mac):* MIAI skal have **Fuld diskadgang** under Systemindstillinger → Anonymitet & sikkerhed. Appen tilbyder at åbne stedet.
- *Modellen svarer i MIAI App's menu, men Office kan ikke nå MIAI:* MIAI har fået en ny port (fx fordi den gamle udgave ikke var lukket helt) og kunne ikke opdatere sig selv i Office. Vælg **Installer i Office…** i menuen, og genstart Word, Excel og PowerPoint. Afvises det, mangler MIAI Fuld diskadgang (se næste punkt). Fra 0.1.11 venter MIAI på sin gamle port og siger til, hvis Office ikke kunne opdateres.
- *Efter en opdatering på Mac* skal Fuld diskadgang gives igen: fjern den gamle MIAI fra listen med **−**, og tilføj den nye fra Programmer med **+**. Nøgleringen kan også spørge igen; vælg *Tillad altid*. Det skyldes, at appen endnu ikke er signeret med et Apple-certifikat, så macOS ser hver udgave som en ny app ([ADR-0016](adr/0016-distribution-signering-opdatering.md)).
- *Loggen* ligger i `~/Library/Application Support/MIAI/MIAI.log` (Mac) og `%APPDATA%\MIAI\MIAI.log` (Windows). Indstillingerne ligger samme sted i `config.json`.

Alt kører lokalt: der er ingen sky-tjeneste og ingen CORS-opsætning, for add-in og bridge har samme adresse.

### B. Hostet add-in (uden app)

Add-in'et ligger på GitHub Pages og kalder selv dit inference-værktøj. Du skal bruge inference-værktøjet og ét manifest.

**1. Hent manifestet:** [`https://mor10w.github.io/MIAIPlugin/manifest.xml`](https://mor10w.github.io/MIAIPlugin/manifest.xml) (højreklik → *Gem link som*). I en fork er adressen `https://<bruger>.github.io/<repo>/manifest.xml`.

**2. Læg det ind i Office**

- **Mac** (Terminal; antager, at filen ligger i Overførsler):
  ```bash
  for app in Word Excel Powerpoint; do mkdir -p ~/Library/Containers/com.microsoft.$app/Data/Documents/wef && cp ~/Downloads/manifest.xml ~/Library/Containers/com.microsoft.$app/Data/Documents/wef/miai.xml; done
  ```
- **Windows** (PowerShell). Læg først filen et sted, hvor den kan blive liggende, fx i en mappe `MIAI` i din brugermappe; Office læser den derfra ved hver start:
  ```powershell
  $m = "$env:USERPROFILE\MIAI\manifest.xml"; $id = ([xml](Get-Content $m)).OfficeApp.Id; New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Force | Out-Null; Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name $id -Value $m
  ```
- Har du hentet repoet, gør `npm run sideload:mac -- <sti til manifest.xml>` og `npm run sideload:windows -- <sti til manifest.xml>` det samme.

**3. Tillad CORS** fra `https://mor10w.github.io` i dit værktøj. Uden det kan add-in'et ikke nå modellen:

- Ollama, Windows: sæt brugermiljøvariablen `OLLAMA_ORIGINS=https://mor10w.github.io` og genstart Ollama.
- Ollama, Mac: `launchctl setenv OLLAMA_ORIGINS https://mor10w.github.io` og genstart Ollama.
- LM Studio: slå "Enable CORS" til under Developer/Server.

**4. Genstart Office.** Knappen **AI-assistent** ligger på fanen **Hjem**.

Brug aldrig `*` som tilladt adresse: det ville lade enhver webside bruge din lokale model. I hostet tilstand ligger indstillinger og eventuelle API-nøgler i add-in'ets lokale lager på maskinen; brug MIAI App, hvis din endpoint kræver en nøgle.

### C. Udvikler-tilstand (fra kildekoden)

```bash
npm install
```

```bash
npm run certs
```

Opretter et udviklingscertifikat for `localhost`; CA'en skal godkendes første gang.

```bash
npm run dev
```

Starter Vite på `https://localhost:3000` med hot reload. Læg derefter udviklingsmanifestet ind i Office (én gang):

```bash
npm run sideload:mac
```

```bash
npm run sideload:windows
```

Genstart Office. Ollama tillader `https://localhost:*` som standard, så der er ingen CORS-opsætning. Mere i [Udvikling og afprøvning](UDVIKLING.md).

### D. MIAI Bridge fra kildekoden

Bridgen er den lokale service, som MIAI App pakker ind. Den serverer add-in'et og kalder modellen, og den giver låst organisationspolitik, audit-log og brug uden internet. Kør den fra kildekoden, hvis du udvikler på den eller vil køre den på en maskine uden skrivebord.

```bash
npm run build
```

```bash
npm start
```

Bridgen kører på `https://localhost:44321` (eller næste ledige port) og skriver dit **personlige manifest** med din port og din adgangsnøgle. Stien vises ved start, fx `~/Library/Application Support/MIAI/manifest.xml`. Læg den fil ind i Office:

```bash
npm run sideload:mac -- ~/Library/Application\ Support/MIAI/manifest.xml
```

På Windows: `npm run sideload:windows -- %APPDATA%\MIAI\manifest.xml`.

Adgangsnøglen betyder, at kun du kan bruge din bridge, også på en delt maskine (Remote Desktop/Citrix), hvor andre brugere kan nå `127.0.0.1`. MIAI App gør alt dette automatisk.

Watch-tilstand for både add-in og bridge: `npm run dev:bridge`.

Tilstanden (direkte eller bridge) vælges automatisk ved start; task pane viser den i overskriften.

## Første gang

1. Åbn Word, Excel eller PowerPoint, og klik **AI-assistent** på fanen **Hjem** (eller Indsæt → Tilføjelser; add-in'et hedder *MIAI*).
2. Task pane åbner til højre med fanerne **Handlinger**, **Chat** og **Indstillinger**.
3. Tjek statusbadget øverst til højre:

| Badge | Betydning |
|---|---|
| 🟢 *Forbundet* | Modellen svarer. Hover viser den aktive endpoint. |
| 🟡 *Ingen forbindelse til modellen* | Endpointet svarer ikke – typisk CORS. Hover viser fejl + vejledning. |
| 🔴 *MIAI Bridge svarer ikke* | Bridgen kører ikke. Start MIAI App, eller `npm start`, hvis du kører bridgen fra kildekoden. |

Ved første start er endpointet **Ollama** (`http://127.0.0.1:11434/v1`) og standardmodellen er den første fra `/v1/models`.

## Indstillinger

Fanen **Indstillinger** (indstillinger gemmes per maskine; i bridge-tilstand af bridgen):

**Førstegangsguiden:** uden forbindelse til en model viser MIAI **Kom i gang**, der åbner Indstillinger med **Find lokale værktøjer**.
- **Søgning:** guiden leder efter Ollama, LM Studio, llama.cpp, vLLM, LocalAI, Jan og KoboldCpp på deres standardporte.
- **Klar:** et værktøj, der svarer, tages i brug med **Brug**.
- **Mangler CORS:** kører værktøjet, men mangler CORS-opsætning (kun direkte tilstand), vises den præcise opsætning med add-in'ets adresse, og kommandoer kan kopieres. MIAI anbefaler aldrig `*`, da det ville lade enhver webside bruge din lokale model.
- **Gem** tester den aktive endpoint først. Svarer den ikke, skal du bekræfte med **Gem alligevel**.

1. **Endpoints** – ét endpoint pr. kort. Vælg en skabelon og klik *Tilføj endpoint*:

   | Skabelon | Standard-URL |
   |---|---|
   | Ollama | `http://127.0.0.1:11434/v1` |
   | LM Studio | `http://127.0.0.1:1234/v1` |
   | llama.cpp server | `http://127.0.0.1:8080/v1` |
   | vLLM | `http://127.0.0.1:8000/v1` |
   | LocalAI | `http://127.0.0.1:8080/v1` |
   | Jan | `http://127.0.0.1:1337/v1` |
   | KoboldCpp | `http://127.0.0.1:5001/v1` |
   | Andet / Other | `http://127.0.0.1:8000/v1` |

   Pr. endpoint kan du sætte navn, basis-URL (inkl. `/v1`), API-nøgle (valgfri), **struktureret output** (`json_schema` / `json_object` / `none`) og *Skjul <think>-tankegang*.
2. **Test forbindelse** – gemmer indstillingerne og henter modellenavnene; viser fx "3 modeller fundet".
3. **Standardmodel** – vælg fra de fundne modeller, eller lad den første tilgængelige bruges.
4. **Model pr. opgave** – hver opgave kan få sin egen model, fx en lille og hurtig til grammatik og oversættelse og en større til formler, diagrammer og dataanalyse. Opgaver uden valg bruger standardmodellen. Findes en valgt model ikke på den aktive endpoint (fx efter skift af værktøj), vises en advarsel.
5. **Temperatur** (0–2) og **Maks. tokens i svar**. Samme sted:
   - **Tænkning** – modeller som Qwen3, DeepSeek-R1 og gpt-oss kan tænke, før de svarer. **Fra** giver markant hurtigere svar, **Til** grundigere. **Serverens standard** sender intet, så modellen gør, som den plejer.
   - **Indsats** – hvor meget modellen tænker (lav, middel, høj). Ikke alle modeller kan styres; de øvrige ignorerer valget. Kan ikke vælges, når tænkningen er slået fra.
   - Afviser serveren eller modellen valget, laver MIAI svaret uden det og skriver det under svaret. Vælg så **Serverens standard**.
6. **Sprog** – Dansk eller English (følger som udgangspunkt Office-sproget).
7. **Web-søgning** (fra som standard) – gør det muligt for modellen at søge online efter aktuelle oplysninger (fx nyheder, versioner, priser). Søgningen kører i bridgen/appen uden API-nøgle, og modellen citerer sine kilder. **Bemærk:** søgeordene dannes af modellen ud fra dit dokument og sendes til DuckDuckGo – det er det eneste, der forlader maskinen, og derfor et aktivt valg. Kræver bridge- eller app-tilstand. En organisation kan forbyde web-søgning i maskinpolitikken (`"webSearch": false`), og med audit-log slået til logges hver søgning (host og tid, aldrig søgeord). **Søgetjeneste:** DuckDuckGo (standard) eller din organisations egen [SearXNG](https://github.com/searxng/searxng), så søgeord ikke sendes direkte til en tredjepart. SearXNG skal have JSON-formatet slået til (`search.formats: [html, json]` i `settings.yml`), og hosten skal stå under Tilladte hosts. Organisationen kan låse søgetjenesten i maskinpolitikken.
8. **Audit-log** – metadata (host, model, tokens) uden indhold; giver mest mening med bridge. Slås til/fra i fanen, og **Vis log** viser de seneste poster (bridge: `audit.log` i konfigurationsmappen, maks. 1 MB med rotation til `.old`; direkte tilstand: de seneste 200 i add-in'et).
9. **Tilladte hosts** – politikken, én host pr. linje. Standard tillader kun loopback; `*.domæne`-wildcards kan bruges. I bridge-tilstand med maskinpolitik er feltet låst: *Politikken er fastsat af din organisation.*
10. **Mappe med brugeropgaver** (kun bridge) – mappen med opgaver fra [Brugerdefinerede opgaver](#brugerdefinerede-opgaver); tom = `<konfig>/tasks`.

Klik **Gem** – uden dette virker ændringerne ikke.

## Brug en intern server

Som standard må MIAI kun kalde modeller på din egen maskine. Har I en server i huset, fx en GPU-maskine med vLLM eller Ollama, kan MIAI bruge den i stedet. Der skal ikke ændres i koden; serveren skal blot tillades.

**Sådan gør du**

1. Åbn fanen **Indstillinger** i MIAI.
2. Tilføj serveren under **Tilladte hosts**, fx `llm.firma.local`. Der står én host pr. linje, og `*.firma.local` tillader alle maskiner under domænet. Lad `127.0.0.1` og `localhost` blive stående, hvis du også vil bruge en lokal model.
3. Tilføj et endpoint (skabelonen *Andet*) med serverens adresse, fx `https://llm.firma.local/v1`, og en API-nøgle, hvis serveren kræver det.
4. Vælg **Gem**. MIAI genindlæser sig selv, så den nye host kan kaldes. Vælg derefter **Test forbindelse**.

**Hvad serveren skal kunne** afhænger af, hvordan MIAI er installeret:

| | MIAI App eller bridge | Hostet add-in |
|---|---|---|
| Hvem kalder serveren | Appen på din maskine | Office selv (task pane) |
| `http` uden kryptering | Virker | Blokeres af browseren; serveren skal køre `https` med et certifikat, maskinen stoler på |
| CORS på serveren | Ikke nødvendig | Skal tillade add-in'ets adresse, fx `https://mor10w.github.io` |
| API-nøgle | Gemmes i systemets nøglering | Gemmes i add-in'ets lokale lager i klartekst |

Appen er derfor den nemmeste vej til en intern server. Hostet kræver mere af serveren, men intet installeret hos brugeren.

På serveren:

- **Ollama** lytter som standard kun på sin egen maskine. Sæt `OLLAMA_HOST=0.0.0.0`, så den kan nås fra netværket, og til det hostede add-in også `OLLAMA_ORIGINS=https://mor10w.github.io`.
- **vLLM:** `--host 0.0.0.0`, og til det hostede add-in `--allowed-origins '["https://mor10w.github.io"]'`.
- Brug aldrig `*` som tilladt adresse. Så kan enhver webside, en bruger besøger, kalde jeres model.

**Lås det for hele organisationen**

En organisation kan lægge serveren ind centralt og låse listen, så brugerne ikke kan tilføje andre:

- **Hostet add-in:** `miai-config.json` ved siden af add-in'et; se [Central udrulning](#central-udrulning-i-en-organisation-også-office-på-web).
- **MIAI App og bridge:** en politikfil på maskinen, som brugeren ikke kan ændre i Indstillinger:
  ```json
  { "allowedHosts": ["llm.firma.local"], "locked": true }
  ```
  Den ligger i `%ProgramData%\MIAI\policy.json` (Windows) eller `/Library/Application Support/MIAI/policy.json` (Mac).

**Hvad det betyder for dine data**

Med en intern server forlader dokumentteksten din maskine, men ikke jeres netværk. MIAI sender kun til de hosts, der står på listen. Websøgning er en indstilling for sig og er slået fra som standard.

## Funktioner

Almindeligt for alle opgaver: markér tekst (hvis det kræves) → klik en opgave i **Handlinger** → evt. vælg parametre → **Kør**. Resultatet vises i task pane, og du vælger selv, hvordan det anvendes – intet sker i dokumentet, før du klikker. **Kopiér** er altid tilgængeligt, og mens der genereres kan du afbryde med **Stop**.

| Opgave | Word | Excel | PowerPoint |
|---|:-:|:-:|:-:|
| Ret stavning og grammatik | ✅ | | ✅ |
| Omskriv i anden tone · Parafrasér · Forenkl | ✅ | | ✅ |
| Oversæt | ✅ | ✅ | ✅ |
| Opsummér | ✅ | | ✅ |
| Skriv artikel ud fra idé | ✅ | | |
| Salgstekst | ✅ | | ✅ |
| Præsentationsdisposition | ✅ | | ✅ |
| Tekst til slides | | | ✅ |
| Lav diagram | ✅ | | ✅ |
| Forklar formel · Lav formel ud fra beskrivelse | | ✅ | |
| Analyser data (tabel) | | ✅ | |
| Chat med dokumentet | ✅ | ✅ | ✅ |

### Word

- **Ret stavning og grammatik** – markér afsnit/tekst. Hver rettelse vises som *gennemstreget original → ny tekst* med kort begrundelse. Vælg **Indsæt som sporede ændringer** (kan accepteres/afvises i Word) eller **Erstat**. Ingen fejl → "Ingen fejl fundet."
- **Omskriv i anden tone / Parafrasér / Forenkl** – markering kræves. Tonen vælges (professionel, venlig, kort og præcis, overbevisende, akademisk, uformel).
- **Oversæt** – markering + målsprog (da/en/de/sv/no/fr/es).
- **Opsummér** – markering, eller uden markering: hele dokumentet. Længde: punktform (3-5 punkter), ét afsnit eller én sætning. Indsættes **efter** markeringen eller **som kommentar**.
- **Skriv artikel ud fra idé** – skriv en linje (fx "5 grunde til at køre AI lokalt"), vælg tone og længde (~300/700/1500 ord). Markér evt. et udkast, den bygges videre på. Indsættes som formateret tekst med overskrifter.
- **Salgstekst** – beskriv produktet eller budskabet, vælg tone (overbevisende/professionel/venlig/drægtig) og længde (~100/250 ord). Markér evt. eksisterende tekst (specifikationer, notater), den bruges som udgangspunkt. Modellen skriver overbevisende tekst med konkret appel og slutter med en opfordring til handling.
- **Præsentationsdisposition** – emne + antal slides (5/8/10/15).
- **Lav diagram** – markér tekst (fx. en tabel eller disposition) → modellen omdanner den til et flydediagram/hierarki (3-10 knuder). Forhåndsvisningen vises i task pane; **Indsæt diagram** indsætter det som billede i et nyt afsnit efter markeringen.

### Excel

- **Forklar formel** – klik en celle med formel. Forklaringen vises i task pane og kan indsættes **som kommentar** på cellen.
- **Lav formel ud fra beskrivelse** – skriv fx "summen af kolonne B hvor kolonne A er 'Jylland'". Modellen foreslår en formel (engelske funktionsnavne, moderne funktioner som XLOOKUP/FILTER prioriteres), den skrives i cellen og **valideres** – fejl som `#NAME?` meldes. Klik **Fortryd** for at gendanne den tidligere formel (Ctrl+Z fortryder ikke altid add-in-ændringer).
- **Oversæt** – markér et område; oversættelsen skrives tilbage i cellerne (antal rækker/kolonner bevares, tal og datoer roterer). Konteksten er max. 200 rækker × 30 kolonner.
- **Analyser data** – markér en tabel (eller brug hele arket). Modellen beskriver, hvad data viser, angiver nøgletal ud fra kolonnens statistikker, og foreslår konkret: en pivot-tabellayout (rækker/kolonner/værdier), diagramtype med præcist område og evt. formler, der verificerer nøgletallene. Resultatet kan indsættes **som kommentar** på det markerede område.

### PowerPoint

- **Omskriv / Parafrasér / Forenkl** – markér tekst i en tekstboks; resultatet erstatter teksten i boksen.
- **Tekst til slides** – indsæt tekst eller beskriv et emne, vælg antal slides (3/5/8/10). Modellen foreslår titler, punkter og taler-notes; **Opret slides** indsætter dem i præsentationen (pladsholdere udfyldes; uden pladsholdere oprettes tekstbokse).
- **Lav diagram** – markér tekst (fx. en tabel) → modellen omdanner den til et flydediagram/hierarki (3-10 knuder). Forhåndsvisningen vises i task pane; **Indsæt diagram** placerer diagrammet som **native figurer** på den valgte slide (afrundede rektangler, linjer og pile, samlet i én gruppe i nyere PowerPoint). Opløs gruppen for at flytte, farve og redigere hver del.
- **Opsummér** – markering, eller hele præsentationen (slide for slide).
- **Præsentationsdisposition** – som i Word.
- **Design: master og layout** – MIAI læser præsentationens indhold, eller dit emne hvis den er tom, og designer én sammenhængende designskabelon til hele præsentationen:
  - **Stil:** ren, markant eller redaktionel.
  - **Farver:** en palet valgt ud fra emnet og publikummet, tjekket for læsbar kontrast.
  - **Skrifttyper:** fra Office' standardsæt.
  - **Layout pr. slide:** titel, sektion, punkter, to kolonner, citat, nøgletal og afslutning.

  Skriv gerne ønsker, fx *"rolig til bestyrelsen, vores farve er #0A5C8A"*. Før noget ændres, ser du navn, begrundelse, farver og en miniature af hvert slide. **Opret** sætter en rigtig slidemaster med navngivne layouts (fx "Nordisk klar – Punkter") og slides med indholdet ind i præsentationen.
  - **Erstat de nuværende slides** fjerner de gamle bagefter (Microsoft 365). Billeder, diagrammer og noter på de gamle slides flyttes ikke med, så kontrollér resultatet. Med **Fortryd** kommer de gamle slides tilbage med al formatering.
  - Layoutene ligger derefter under **Nyt slide**, og agenten bruger dem, når den laver nye slides, så præsentationen bliver ved med at følge designet.

### Outlook

MIAI virker i Outlook (Windows, Mac, nyt Outlook og Outlook på web) med et eget manifest.

**Installation:** Hent tilføjelsesprogrammer → Mine tilføjelsesprogrammer → Tilføj brugerdefineret → Tilføj fra fil. Brug en af disse filer:
- **Hostet:** [`https://mor10w.github.io/MIAIPlugin/manifest-outlook.xml`](https://mor10w.github.io/MIAIPlugin/manifest-outlook.xml) (i en fork: `https://<bruger>.github.io/<repo>/manifest-outlook.xml`).
- **MIAI App:** menuen → **Vis Outlook-manifest i Finder…** (Windows: *Stifinder*).
- **Bridgen:** `manifest-outlook.xml` i konfigurationsmappen; stien vises, når bridgen starter.

Knappen **AI-assistent** findes derefter i mailen og i udkastet.
- **Når du læser en mail:**
  - **Svar på mail**: fortæl evt., hvad svaret skal sige, fx *"takker ja, men foreslår torsdag"*. Svaret åbnes som et udkast (**Svar** eller **Svar alle**), som du retter og sender selv. MIAI sender aldrig noget. Mangler der oplysninger, skriver MIAI `[pladsholdere]` i stedet for at opfinde dem.
  - **Opsummér**, **Oversæt** og **chat** om mailen, fx *"Er der frister eller åbne spørgsmål?"*. Emne og afsender er med som kontekst.
- **Når du skriver en mail:** **Omskriv i anden tone**, **Ret stavning og grammatik**, **Forenkl**, **Oversæt** og **Salgstekst** virker på markeringen eller hele udkastet. **Erstat** indsætter resultatet med formatering (HTML) i udkastet. Outlooks egen fortryd (Ctrl+Z) virker bagefter.

### Chat (alle tre apps)

Spørg om dit dokument eller få hjælp til Office. **Brug markering/dokument som kontekst** (til som udgangspunkt) giver modellen din aktuelle markering – i Excel dit markerede område – med. De seneste 20 beskeder medsendes som historik. Svar kan indsættes **efter** markeringen, **erstattes** eller indsættes **som kommentar**. **Ny samtale** nulstiller historikken. Enter sender, Shift+Enter linjeskift.

- **Vedhæft billeder og filer**: klik på klipsen ved skrivefeltet, eller sæt et skærmbillede ind med Ctrl+V.
  - **Billeder** (PNG, JPEG, WebP, GIF) sendes til modellen sammen med beskeden. Det kræver en model, der kan se (fx gemma3 eller qwen2.5vl). Kan modellen ikke det, får du en besked om at vælge en anden model eller fjerne billedet. Store billeder skaleres ned først.
  - **Tekstfiler** (.txt, .md, .csv, .json m.fl.) læses som tekst. **PDF, Word, Excel og PowerPoint** kan vedhæftes med MIAI Bridge eller MIAI App, som trækker teksten ud; filen gemmes ikke.
  - Højst 6 vedhæftninger pr. besked og 60.000 tegn pr. fil (resten udelades, og modellen får det at vide).
  - Vedhæftninger går også med, når chatten laver en plan for dokumentet, fx *"indsæt tabellen fra billedet i arket"*.
  - I opfølgende beskeder husker modellen tekstfilernes indhold, men billedet sendes ikke igen.
- **Markeringen vises**: under skrivefeltet står, hvad chatten bruger som kontekst lige nu, fx "Markering: Ark1!B2:B9", "Markering: 42 ord" eller "Intet markeret – hele dokumentet bruges". Den følger med, når du markerer noget andet.
- **Samtalen bliver stående, når du skifter fane**, fx til Indstillinger og tilbage: et svar, der er ved at blive skrevet, skrives færdigt imens, og en plan, din kladde og dine vedhæftninger er der stadig.
- **Samtalen huskes pr. dokument**: åbner du MIAI igen på samme dokument, er samtalen der stadig, og den går med som historik til modellen.
  - Den gemmes kun på denne maskine (i task panes lager), højst 40 beskeder pr. dokument og 20 dokumenter.
  - Dokumenter, der ikke er gemt endnu, huskes ikke.
  - Planer fra agenten gemmes som tekst (opsummering og udfald) og kan ikke udføres igen fra en gammel samtale.
  - **Ny samtale** sletter den. Under *Indstillinger* kan du slå **Husk samtalen pr. dokument** fra.
- **Klikbare henvisninger**: i Excel bliver celleadresser i et svar (`C2:C5`, `B7`, `Budget!B2`) til links, der markerer cellerne og skifter til arket; i PowerPoint går "slide 3" til sliden. En enkelt celle linkes kun, når den ligger inden for arkets data, så "Q1 2026" ikke bliver et link. Adresser i kodeblokke linkes ikke.
- **Figurer i chatten**: Hvis svaret indeholder et diagram (` ```mermaid `) eller en selvstændig SVG (```svg), vises det som billede i task pane, og knappen **Indsæt som billede** indsætter det – i Word som inline-billede ved markeringen, i PowerPoint som billedfyldt rektangel på den valgte slide (PowerPointApi 1.8). Knappen finder også figurer, modellen har gærdat sløjt (fx SVG under ` ```xml `, en afbrudt blok eller rå `<svg>`-markup uden gærde). Typiske diagramfejl fra små modeller (etiketter med parenteser uden anførselstegn, `<b>`/`<i>` i etiketterne) rettes automatisk. Kan diagrammet stadig ikke tegnes, står fejlen under svaret, og **Bed modellen rette diagrammet** sender den til modellen.
- **Web-søgning**: Modellen kan selv søge online, når den mangler aktuelle oplysninger (indstillingen *Web-søgning*, fra som udgangspunkt). Undervejs vises "Søger på webbet: …", og svaret citerer sine kilder. Søgeresultater er utroværdig input og behandles som data – kritiske fakta bør verificeres.

### Agent (Word, Excel og PowerPoint)

**Agent: redigér dokumentet** udfører ændringer for dig. Beskriv, hvad der skal ændres, fx *"ret alle 'kunde' til 'klient' og tilføj overskriften 'Vilkår' øverst"* eller i Excel *"tilføj en total-række under tabellen"*.

1. **Plan.** MIAI foreslår en liste af konkrete handlinger med en kort begrundelse for hver. **Intet er ændret endnu.**
   - Word: erstat tekst, indsæt afsnit, skift et afsnits typografi (fx brødtekst → overskrift), slet eller flyt hele afsnit, kommentér, indsæt en tabel og tilføj eller slet tabelrækker.
   - Excel:
     - **Data:** skriv værdier, skriv en formel, tilføj eller slet rækker (også *"fjern de tomme rækker"*), ryd et område og sortér.
     - **Udseende:** giv en tabel et moderne udseende med en af Excels tabeltypografier, sæt talformat, tilpas kolonnebredden og brug betinget formatering (farveskala, datalinjer, fremhæv over/under en værdi, dubletter, top 10).
     - **Visning:** lav et diagram og frys overskriftsrækken.
     - **Analyse og indtastning:** lav en pivottabel (fx *"omsætning pr. region"*, gerne på et nyt ark) og sæt datavalidering (en rulleliste med tilladte værdier eller et tal mellem to grænser).
     - **Flere ark:** agenten kender alle ark i projektmappen (navn, brugt område og overskrifter) og kan skrive på et andet ark, oprette et nyt ark og omdøbe et ark, fx *"lav et nyt ark med den samlede omsætning"*. Fortryd dækker alle berørte ark. Værdierne ser den kun for markeringen.
   - PowerPoint: erstat tekst, skriv ny tekst i en figur, opret et slide fra præsentationens layouts, slet og flyt slides.
   - I Excel kan agenten først **læse** det, den mangler at se, fx et andet ark eller et område uden for markeringen. Det ændrer intet og kræver ingen godkendelse; du ser "Læser Budget!A1:F40…", mens det sker. Den kan læse højst tre gange, før den skal lave planen.
2. **Bekræft.** Fjern fluebenet ved de handlinger, du ikke vil have udført. **Vis i dokumentet** under en handling markerer det sted, den gælder (celler, tabel, afsnit, tekst eller slide), uden at ændre noget.
   - Handlinger, der sletter indhold, er markeret *Fjerner indhold*.
   - Handlinger, der ikke kan udføres, er fravalgt med årsagen. Det gælder fx tekst, modellen har citeret forkert, så den ikke findes i dokumentet.
   - I Word udføres ændringerne som standard **som sporede ændringer**, så du også kan gennemgå dem i Word.
3. **Udfør.** Hver handling får ✓ eller en forklaring. Fejler nogle af handlingerne, kan du trykke **Bed modellen rette det, der fejlede**: fejlene sendes til modellen, som foreslår en rettelse. Det, der lykkedes, er udført og gentages ikke.

**Formateringen bevares.** Agenten ser dokumentet afsnit for afsnit med typografi (overskrift, brødtekst, punkt, tabel), så den kan arbejde med strukturen frem for imod den:
- **Erstat** ændrer kun de ord, der faktisk er forskellige. Fed, kursiv og links omkring dem bliver stående.
- **Ny tekst** får sin egen typografi og arver ikke naboens formatering. En linje efter en overskrift bliver altså brødtekst, ikke endnu en overskrift. Agenten kan selv skabe formatering: overskrifter, punkter, nummererede lister, citater, **fed**, *kursiv* og links. Skrifttype, farve og størrelse følger dokumentets egne typografier.
- **Skift af typografi** sker uden at røre teksten.
- **Slet og flyt** virker på hele afsnit. Flyttede afsnit beholder deres formatering, og slettede efterlader ingen tomme linjer.
- **Tabeller:** agenten ser hver celle med tabel, række og kolonne. Den kan rette teksten i en celle, tilføje en række (*"tilføj Pære til 5 kr i pristabellen"*), slette en række og indsætte en ny tabel. En ny tabel får samme tabeltypografi som dokumentets første tabel; har dokumentet ingen, bruges en indbygget. Agenten flytter ikke celler og indsætter ikke almindelige afsnit i en tabel.
- **Tabeller fra chatten:** en Markdown-tabel i et chatsvar vises som tabel og bliver en rigtig Word-tabel, når du vælger *Indsæt efter* eller *Erstat*.
- Har du ændret dokumentet, efter planen blev lavet, rører agenten ikke de berørte afsnit. Kør den igen.

**I Excel** bevares formateringen også:
- Værdier og formler ændrer aldrig cellernes format.
- Tomme celler, der fyldes ud lige under eller ved siden af data, får naboens format. En ny total-række får altså kolonnens valuta- eller procentformat.
- Agenten kan se arkets Excel-tabeller og tilføjer rækker *i* tabellen, så de får tabellens formatering og beregnede kolonner. Slettes rækker i en tabel, krymper tabellen; data under den rykker op.
- *"Gør tabellen pæn"* giver tabellen en af Excels indbyggede tabeltypografier, som følger projektmappens tema. Et almindeligt område bliver til en Excel-tabel med filterknapper og stribede rækker. Agenten vælger aldrig selv farver eller skrifttyper.

**Kodetilstand (avanceret, i Excel, Word og PowerPoint).** Under *Indstillinger* kan du lade agenten skrive og køre Office.js-kode. Det er til det, de faste handlinger ikke kan, fx navngivne områder, sortering på flere kolonner og filtre i pivottabeller.
- **Før kørslen:** du ser, hvad scriptet gør, og kan se koden.
- **Sandkassen:** scriptet kører uden adgang til netværk, filer eller MIAI's indstillinger.
- **Hvis det fejler:** arket sættes straks tilbage, og du kan bede modellen rette scriptet ud fra fejlen.
- **Fortryd i Excel:** gendanner det aktive ark og fjerner nye ark, tabeller, diagrammer og pivottabeller. Betinget formatering, datavalidering og ændringer på andre ark gendannes ikke.
- **Fortryd i Word:** sætter dokumentets brødtekst tilbage til, som den var før scriptet – også ændringer, du selv har lavet siden. Sidehoved, sidefod og kommentarer gendannes ikke.
- **Fortryd i PowerPoint:** lægger alle slides tilbage, som de var før scriptet. Det kræver Microsoft 365 (PowerPointApi 1.8); ellers køres scripts ikke. Slidemasteren gendannes ikke.
- **Begrænsninger:** i Excel virker kodetilstanden kun på ark med højst 50.000 celler. Den kræver en stærk model (14B eller større). Din organisation kan have slået den fra.

**Chatten kan selv ændre dokumentet.** Beder du i chatten om en ændring (*"fjern de tomme rækker"*, *"gør overskriften kortere"*), laver MIAI en plan direkte i samtalen. Du godkender handlingerne, og planen kan fortrydes, præcis som under *Agent: redigér dokumentet*.
- **Opfølgninger virker:** samtalen, de tidligere planer og hvad der faktisk blev udført, går med til modellen. Du kan altså skrive *"nej, sortér faldende i stedet"*.
- **Spørgsmål besvares som før** med tekst. Kontakten **Må ændre dokumentet** under skrivefeltet slår ændringer fra, hvis chatten kun skal svare.
- **Hvis MIAI ikke opfatter beskeden som et ønske om en ændring,** svarer den med tekst. Henviser svaret til agenten, vises knappen **Gør det med agenten**, som laver planen i samtalen.
- MIAI kan ikke køre VBA eller makroer.

**I PowerPoint** følger agenten skabelonen:
- Erstat ændrer kun de ord, der er forskellige.
- Ny tekst i en figur beholder figurens skrifttype, størrelse, farve og punkttegn.
- Nye slides oprettes fra præsentationens egne layouts med teksten i pladsholderne. Agenten laver aldrig egne tekstbokse, farver eller skrifttyper.

Hele sessionen kan fortrydes med **Fortryd agent-session**:
- I Word med sporede ændringer (standard, Word 2021/Microsoft 365) afvises agentens ændringer og kommentarer, og resten af dokumentet røres ikke.
- I ældre Word eller uden sporede ændringer gendannes indholdet fra lige før. Det fjerner også ændringer, du selv har lavet siden.
- I Excel gendannes de berørte celler med deres formatering, og tilføjede tabelrækker fjernes.
- I PowerPoint (Microsoft 365) lægges de berørte slides tilbage med al formatering, nye slides slettes, og flyttede flyttes tilbage. Ældre PowerPoint kan kun gendanne teksten, og derfor sletter og flytter agenten ikke slides dér.

Fra **chatten** kan du sende et ændringsønske videre med **Gør det med agenten**. Agenten bruger chatsvaret, du lige har set, som indhold i stedet for at skrive en ny tekst.

Agenten arbejder på markeringen, eller på hele dokumentet hvis intet er markeret. I Excel arbejder den kun på det aktive ark, og i PowerPoint ser den hele præsentationen. Den har ingen adgang til netværk eller filer, og instruktioner, der står i selve dokumentet, følges ikke. Se [ADR-0012](adr/0012-agent-tilstand.md).

### Viden fra mine mapper (bridge og MIAI App)

Chatten kan svare ud fra dine egne filer og vise, hvilke filer svaret bygger på. Filerne bliver på maskinen. Kun de relevante tekstuddrag sendes til din egen model.

1. **Hent en embedding-model** i dit inference-værktøj. Det er en lille model, der gør tekst søgbar:

   | Model | Hvornår | Hent (Ollama) |
   |---|---|---|
   | **bge-m3** (anbefalet) | Dansk og blandede sprog, god kvalitet | `ollama pull bge-m3` |
   | embeddinggemma | Lettere, flersproget, til laptops | `ollama pull embeddinggemma` |
   | nomic-embed-text | Hurtig, mest til engelsk | `ollama pull nomic-embed-text` |
   | qwen3-embedding | Organisationer med en GPU-server | `ollama pull qwen3-embedding` |

   LM Studio, llama.cpp og vLLM virker også, hvis de understøtter `/v1/embeddings`. Gør de ikke, siger MIAI det.
2. **Indstillinger → Viden fra mine mapper.** Slå *Søg i mine mapper* til, og tilføj mapper, fx `C:\Users\dig\Documents\Politikker` eller `~/Dokumenter`. Vælg embedding-modellen og tryk **Gem**. Indekseringen starter, og status viser antal filer og uddrag.
3. **I chatten** er *Søg i mine mapper* slået til. Svaret citerer kilderne som [1], [2] …, og under svaret står **Kilder fra dine mapper**. Klik på en kilde for at se uddraget og stien. **Åbn fil** åbner filen i Word, Acrobat osv.

**Hvad der indekseres:**
- Filtyper: Word (.docx), PDF, Excel (.xlsx), PowerPoint (.pptx), Markdown, tekst og CSV.
- Springes over: skjulte filer, Office' låsefiler (`~$…`) og filer over 30 MB.
- Loftet er 10.000 filer.

**Opdatering:** ændringer i mapperne opdages automatisk og indekseres efter få sekunder, og kun ændrede filer læses igen. *Indeksér nu* tvinger en opdatering. Skifter du embedding-model, bygges indekset forfra.

**Hvor data ligger:** indekset ligger i `knowledge.db` ved siden af `config.json`. Slet filen for at fjerne alt.

**Politik:** en organisation kan slå funktionen fra (`"knowledge": false` i `policy.json`) eller begrænse den til bestemte mapper (`"knowledgeFolders": ["D:\\Fælles\\Håndbøger"]`). Med audit-log slået til logges indeksering og opslag med antal og varighed, aldrig stier, spørgsmål eller indhold. Se [ADR-0013](adr/0013-lokal-rag.md).

## Brugerdefinerede opgaver

Fanen **Opgaver** viser dine egne opgaver, som du kan oprette, redigere, slette og importere/eksportere. En opgave er en JSON-fil:

```json
{
  "id": "resume",
  "label": { "da": "Opsummér til chefen", "en": "Summarize for the boss" },
  "description": { "da": "Kort resumé i tre punkter." },
  "hosts": ["word"],
  "input": "selection",
  "output": { "format": "markdown" },
  "applyModes": ["insertAfter", "comment"],
  "temperature": 0.4,
  "prompt": {
    "system": "Du er en erfaren sekretær. Skriv formelt og kort.",
    "user": "{{selection}}\n\nSammenfat i 3 punkter på {{language}}."
  }
}
```

Skabelonen udfyldes med `{{selection}}` (markeringen), `{{prompt}}`, `{{language}}`, `{{address}}` (celleområde i Excel) og `{{table}}` samt opgavens egne parametre. Ukendte placeholderer efterlates som de er, så du opdager dem. Når skabelonen bruger `{{selection}}`, pakkes markeringen i `<document>`-tags, så den behandles som data.

**Deling:**
- **Med MIAI Bridge** ligger opgaverne i en mappe (indstillingen *Mappe med brugeropgaver*, standard `<konfig>/tasks`). Kopier mappen til en kollega – færdig. Opgaver, IT uddeler via `MIAI_TASKS_DIR`, vises med mærket *Organisation* og kan ikke ændres eller slettes. En bruger-opgave må have samme id som en indbygget opgave, hvormed prompten tilpasses lokalt.
- **Direkte tilstand** (uden bridge) kan browsere ikke læse mapper: brug *Importér fil* og *Eksportér* i Opgaver-fanen. Opgaverne gemmes i add-in'et på maskinen.

Filer kan kun udtrykke tekstoutput (fritekst/Markdown/TSV) – opgaver som grammatik, formler og diagrammer kræver stadig kode.

## Anvendelsesknapper pr. app

| Knap | Word | Excel | PowerPoint |
|---|:-:|:-:|:-:|
| Erstat markeringen | ✅ | ✅ (celleværdi) | ✅ (tekst i boksen) |
| Indsæt som sporede ændringer | ✅ (WordApi 1.4) | | |
| Indsæt efter markeringen | ✅ | | ✅ |
| Indsæt diagram | ✅ (billede) | | ✅ (figurer) |
| Som kommentar | ✅ (WordApi 1.4) | ✅ (ExcelApi 1.10) | |
| Skriv i celler | | ✅ (oversættelse af et markeret område – vises som tabel inden indsættelse) | |
| Indsæt formel (+ Fortryd) | | ✅ (formelgenerering) | |
| Opret slides | | | ✅ (PowerPointApi 1.4) |
| Indsæt som billede (diagram/figur fra chatten) | ✅ | | ✅ (PowerPointApi 1.8) |
| Kopiér | ✅ | ✅ | ✅ |

Knapper, som din Office-version ikke understøtter, skjules.

## Tips

- **Modellens størrelde** afgør kvaliteten. `qwen3:4b` er et pænt udgangspunkt; brug en større model (7B+) til artikler og oversættelser, hvis maskinen rækker.
- **Reasoning-modeller** (fx Qwen3) sender `<think>`-blokke med – indstillingen *Skjul <think>-tankegang* er slået til som standard, så tankegangen ikke ender i dokumentet.
- **Lange dokumenter og kontekstlængde**: hver model har en kontekstgrænse (hvor meget tekst den kan læse ad gangen).
  - **Opsummér** deler lange dokumenter (op til ca. 60 sider) i dele, opsummerer hver del og samler til sidst. Mens det sker, vises "Læser lang tekst: del 3 af 12…".
  - Andre opgaver sender højst 24.000 tegn. Bliver teksten afkortet, siger MIAI det.
  - Har modellen alligevel ikke fået hele teksten med, vises en advarsel. MIAI sammenligner modellens egen optælling med det sendte; det sker fx med Ollama, hvis `num_ctx` er mindre end teksten. Øg så kontekstlængden i værktøjet (Ollama: `OLLAMA_CONTEXT_LENGTH`), og skriv den under **Kontekstlængde** for endpointen.
  - Kendes kontekstlængden (angivet eller aflæst fra fx vLLM), begrænses svarlængden, så prompt og svar passer, og en alt for lang tekst afvises, før den sendes.
  - Chat om meget store dokumenter kommer med lokal RAG (fase 4).
- **Fejlbeskeder** vises på dansk eller engelsk med en forklaring (fx "tjek API-nøglen" ved 401/403) og serverens egen besked som detalje.
- **Timeout**: MIAI giver op, når modellen har været tavs længere end endpointens **Timeout** (Indstillinger → endpointens *Avanceret*; standard 5 min, 10–1800 s). Det gælder både tiden, før svaret begynder (fx mens en lang prompt behandles), og pauser undervejs. Et svar, der skrives løbende, afbrydes aldrig, uanset hvor langt det er. En organisation kan sætte værdien centralt som `timeoutMs` på endpointen i `miai-config.json`.
- **Gentagelsessløjfer**: små, kvantiserede modeller kan gå i stå og gentage den samme frase. MIAI afbryder svaret, når slutningen gentager sig selv mange gange i træk, og viser en advarsel. Teksten forurenes ikke med en note. Vagten gælder kun fritekst; Excel-celler (TSV) og strukturerede svar (JSON) har legitime gentagelser og valideres på anden vis. Sker det ofte, så vælg **Modvirk gentagelser** under Indstillinger for endpointen. Det sender `presence_penalty` til modellen; Qwen anbefaler 1,0–1,5 til kvantiserede modeller. Den bruges kun til fritekst, ikke til JSON, formler eller Excel-celler. En for høj værdi kan give sprogblanding. Alternativt kan du prøve en større model.
- **Struktureret output**: opgaver som grammatik, slides, formler og diagrammer kræver JSON. Værktøjer uden støtte (fx KoboldCpp) sættes til `none`, og validering med genforsøg tager over – det virker, men er mere sårbart for små modeller.
- **Dokumenttekst behandles som data**, aldrig som instruktioner, så markeringer med fx "ignorer ovenstående" ikke styrer modellen.
- **Indstillinger i direkte tilstand** ligger i add-in'ets `localStorage` på maskinen. Nogle webview-bygninger rydder den; i så fald gendannes standardindstillingerne.
- **API-nøgler**: bridgen og MIAI App gemmer nøgler i OS'ets nøglering (Windows Credential Manager, macOS Keychain, Linux Secret Service), og `config.json` indeholder kun en reference (`keyring:<endpoint-id>`). Nøgler fra en tidligere version flyttes automatisk ved første start. Findes der ingen nøglering (fx Linux uden skrivebord), bliver nøglerne i `config.json`, som kun du kan læse, og bridgen skriver en advarsel i loggen. I **direkte tilstand** har browseren ingen nøglering, så nøgler ligger i klartekst i `localStorage`. Brug bridgen eller MIAI App, når din endpoint kræver en nøgle.

## Udrulning til kunder (salgs løsninger)

To salgsveje på samme kodebase – se [ADR-0008](adr/0008-distribution-af-solgte-loesninger.md):

| Udgave | Installation hos kunden | Krav |
|---|---|---|
| **Hostet** (M365) | IT-adminen uploader manifestet i M365-admin center (Integrated apps portal); endbruger ser add-in'et uden opsætning. Web-indhold hostes af leverandøren på eget domæne | M365-admin + `OLLAMA_ORIGINS` for leverandørens domæne |
| **Lokal app** (.dmg/.exe) | Kunden installerer MIAI App (vej A ovenfor); bridgen kører lokalt, manifest sideloades automatisk | Ingen – alt lokalt, ingen internet |

Model-URL og valg af model sker i add-in'ets fanen **Indstillinger** i begge udgaver.

### Central udrulning i en organisation (også Office på web)

Til organisationer, der vil styre MIAI centralt og holde data i huset ([ADR-0015](adr/0015-outlook-og-central-udrulning.md)):

1. **Host add-in'et selv:** læg indholdet af `packages/addin/dist` (fra `npm run build`) på en intern HTTPS-server, fx `https://miai.firma.local/`.
2. **Lav manifester til jeres adresse:**
   ```bash
   node scripts/make-manifest.mjs hosted https://miai.firma.local/ manifest.xml
   ```
   ```bash
   node scripts/make-manifest.mjs hosted https://miai.firma.local/ manifest-outlook.xml outlook
   ```
3. **Læg `miai-config.json` ved siden af `taskpane.html`**, med organisationens politik og standardindstillinger:
   ```json
   {
     "policy": { "allowedHosts": ["llm.firma.local"], "webSearch": false, "knowledge": false },
     "settings": {
       "endpoints": [{ "id": "firma", "name": "Firmaets model", "baseUrl": "https://llm.firma.local/v1" }],
       "activeEndpoint": "firma",
       "auditLog": true
     }
   }
   ```
   - Politikken er altid låst. Brugerne kan ikke ændre den, og Indstillinger viser "Organisationen (central udrulning)".
   - Browseren håndhæver også politikken: add-in'et sætter en Content-Security-Policy, så task pane kun kan sende data til sin egen adresse og de tilladte hosts. Sæt gerne samme CSP som header på jeres webserver (`connect-src 'self' https://llm.firma.local:*; img-src 'self' data: blob:; object-src 'none'; form-action 'none'; base-uri 'self'`).
   - Indstillingerne er standard for alle, der ikke selv har gemt noget.
   - Kan filen ikke læses, låses politikken til standardens hosts, så en fejl i filen aldrig åbner for mere.
4. **Tillad CORS fra add-in'ets adresse på inference-serveren**, fx `OLLAMA_ORIGINS=https://miai.firma.local` eller vLLM `--allowed-origins`.
5. **Udrul i M365 Admin Center** → Indstillinger → Integrated apps → Upload custom apps. Upload `manifest.xml` (Word, Excel, PowerPoint) og `manifest-outlook.xml` (Outlook), og vælg brugere eller grupper. Add-in'et vises så hos brugerne, også i Office på web, uden lokal installation.

## Fejlfinding

| Symptom | Løsning |
|---|---|
| Gul badge / "CORS"-fejl i fejlbeskeden | Værktøjet tillader ikke kald fra add-in'ets adresse. Ollama: `OLLAMA_ORIGINS` (se Installation, vej B). LM Studio: "Enable CORS". vLLM: `--allowed-origins`. Alternativt brug bridgen, som kører uden CORS. |
| "Endpoint er ikke tilladt af politikken." | Endpointet ligger uden for *Tilladte hosts*. Tilføj hosten (fx `llm.firma.local`) og **Gem** – medmindre en maskinpolitik låser feltet. Se [Brug en intern server](#brug-en-intern-server). |
| Rød badge "MIAI Bridge svarer ikke" | Bridgen kører ikke. Start MIAI App (ikonet skal ligge i menulinjen eller ved uret). Kører du bridgen fra kildekoden: `npm start`, og brug dit personlige manifest (stien vises ved start). |
| Ingen modeller i udfyldningsfeltet / "Test forbindelse" fejler | Kører værktøjet på den rigtige port? Prøv `curl http://127.0.0.1:11434/v1/models`. |
| Svarene er langsomme, og modellen "tænker" længe først | Sæt **Tænkning** til **Fra** under Indstillinger → Svar. Det virker med modeller, der kan begge dele (fx Qwen3); rene tænke-modeller tænker altid. |
| "Modellen svarede ikke i tide" | Modellen var tavs længere end endpointens **Timeout** (standard 5 min). Det sker typisk med store dokumenter på langsom hardware, før modellen begynder at svare. Hæv den under Indstillinger → endpointens **Avanceret** → **Timeout** (op til 30 min), eller vælg en mindre model. |
| Formel giver `#NAME?` efter indsættelse | Modellen brugte et funktionsnavn, din Excel ikke kender. Ret formelen manuelt eller klik **Fortryd**. |
| Indstillinger genværrer ikke | Direkte tilstand gemmer i `localStorage`; nogle Mac-bygninger rydder den. Start på ny og gem igen, eller brug bridgen (gemmer i `config.json`). |
| Task pane åbner ikke | Genstart Office efter sideload. Fjern evt. add-in'et (Indsæt → Tilføjelser → …) og sideload igen. |

## Se også

- [Udvikling og afprøvning](UDVIKLING.md) – byg, tests, konfigurationsfiler
- [Arkitektur](ARKITEKTUR.md) – komponenter, dataflow, sikkerhed
- [Projektplan](PLAN.md) – faser og funktionskatalog
- [Arkitekturbeslutninger](adr/) – bl.a. ADR-0006 (politik) og ADR-0007 (bridge vs. direkte)
