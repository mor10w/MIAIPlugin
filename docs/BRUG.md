# Brugervejledning

Sådan installerer og bruger du MIAI – AI-assistenten til Word, Excel og PowerPoint med dine egne lokale modeller.

> Status: fase 0 (spike). Skelettet fungerer, men er ikke fuldt afprøvet i Office. Se [testlisten](UDVIKLING.md#fase-0--manuel-testliste) for, hvad der mangler.

## Hvad du skal bruge

- **Microsoft 365 / Office 2019+ desktop** (Windows eller Mac). Nogle funktioner kræver nyere API'er:
  - *Sporede ændringer* og *kommentarer i Word*: WordApi 1.4
  - *Opret slides* og *opsummér hele præsentationen*: PowerPointApi 1.4
- **Et OpenAI-kompatibelt inference-værktøj** med en model, fx:
  - [Ollama](https://ollama.com): `ollama pull qwen3:4b`
  - LM Studio (start "Local Server")
  - `llama-server -m model.gguf --port 8080`, vLLM, LocalAI, Jan m.fl.
- **Node.js 22 LTS** – kun hvis du kører udvikler- eller bridge-tilstand (hostet tilstand kræver ingen byg).

Dokumentet forlader aldrig din maskine: standardpolitikken tillader kun endpoints på loopback (`127.0.0.1`, `localhost`).

## Installation – vælg én af tre veje

### A. Udvikler-tilstand (direkte) – anbefales til første gang

```bash
npm install
npm run certs    # udviklingscertifikat for localhost, CA'en skal godkendes første gang
npm run dev      # Vite på https://localhost:3000 med hot reload
```

Sideload derefter add-in'et (én gang):

```bash
npm run sideload:mac        # macOS
npm run sideload:windows    # Windows
```

Genstart Office. Ollama tillader `https://localhost:*` som standard, så der er ingen CORS-opsætning.

### B. Med MIAI Bridge (valgfri)

Bridgen er en lokal Node.js-service, der serverer add-in'et og kalder modellen. Den giver låst organisationspolitik, audit-log og UI uden internet.

```bash
npm run build
npm start                # bridgen kører på https://localhost:44321 (eller næste ledige port)
```

Bridgen skriver dit **personlige manifest** med din port og din adgangsnøgle. Stien vises ved start, fx `~/Library/Application Support/MIAI/manifest.xml`. Sideload den fil:

```bash
npm run sideload:mac -- ~/Library/Application\ Support/MIAI/manifest.xml
```
På Windows: `npm run sideload:windows -- %APPDATA%\MIAI\manifest.xml`.

Adgangsnøglen betyder, at kun du kan bruge din bridge, også på en delt maskine (Remote Desktop/Citrix), hvor andre brugere kan nå `127.0.0.1`. MIAI App installerer og opdaterer det personlige manifest automatisk.

Watch-tilstand for både add-in og bridge: `npm run dev:bridge`.

### C. Hostet (uden at bygge noget)

Add-in'et hostes på GitHub Pages, og add-in'et kalder selv dit inference-værktøj. Du skal kun bruge inference-værktøjet og ét manifest:

1. Hent [`https://mor10w.github.io/MIAIPlugin/manifest.xml`](https://mor10w.github.io/MIAIPlugin/manifest.xml) og sideload det (i en fork: `https://<bruger>.github.io/<repo>/manifest.xml`).
2. Tillad CORS fra `https://mor10w.github.io` i dit værktøj:
   - Ollama, Windows: sæt brugermiljøvariablen `OLLAMA_ORIGINS=https://mor10w.github.io` og genstart Ollama.
   - Ollama, Mac: `launchctl setenv OLLAMA_ORIGINS https://mor10w.github.io` og genstart Ollama.
   - LM Studio: slå "Enable CORS" til under Developer/Server.
3. Genstart Office.

Tilstanden (direkte vs. bridge) vælges automatisk ved start; task pane viser den i overskriften.

### D. MIAI App (.dmg) – uden Node.js, til salgs

MIAI App er en menu-bar-app, der kører bridgen lokalt – brugeren behøver hverken Node.js eller at drive en web-side ([ADR-0008](adr/0008-distribution-af-solgte-loesninger.md)):

1. Træk **MIAI** fra .dmg'en til Applications og start appen (første gang: højreklik → Åbn, mens den ikke er signerede).
2. Ved første start genererer appen et TLS-certifikat for `localhost` og tilbyder at betro det og installere MIAI i Office i ét trin (**Gør det nu**). Dialogen vises kun igen, hvis noget mangler, fx når certifikatet fornyes (ca. hvert andet år).
3. Genstart Word/Excel/PowerPoint. Begge trin findes også i menuen (**Betro certifikat…**, **Installer i Office…**).
4. Vælg endpoint (model-URL) og model i add-in'ets fanen **Indstillinger**.

Alt kører lokalt: ingen vendor-cloud, ingen CORS (add-in og bridge er samme origin), indstillinger ligger i `~/Library/Application Support/MIAI/config.json`.

## Første gang

1. Åbn Word, Excel eller PowerPoint → **Indsæt → Tilføjelser** (add-in'et hedder *MIAI*).
2. Task pane åbner til højre med fanerne **Handlinger**, **Chat** og **Indstillinger**.
3. Tjek statusbadget øverst til højre:

| Badge | Betydning |
|---|---|
| 🟢 *Forbundet* | Modellen svarer. Hover viser den aktive endpoint. |
| 🟡 *Ingen forbindelse til modellen* | Endpointet svarer ikke – typisk CORS. Hover viser fejl + vejledning. |
| 🔴 *MIAI Bridge svarer ikke* | Bridgen kører ikke. Start den med `npm start`. |

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
5. **Temperatur** (0–2) og **Maks. tokens i svar**.
6. **Sprog** – Dansk eller English (følger som udgangspunkt Office-sproget).
7. **Web-søgning** (fra som standard) – gør det muligt for modellen at søge online efter aktuelle oplysninger (fx nyheder, versioner, priser). Søgningen kører i bridgen/appen uden API-nøgle, og modellen citerer sine kilder. **Bemærk:** søgeordene dannes af modellen ud fra dit dokument og sendes til DuckDuckGo – det er det eneste, der forlader maskinen, og derfor et aktivt valg. Kræver bridge- eller app-tilstand. En organisation kan forbyde web-søgning i maskinpolitikken (`"webSearch": false`), og med audit-log slået til logges hver søgning (host og tid, aldrig søgeord). **Søgetjeneste:** DuckDuckGo (standard) eller din organisations egen [SearXNG](https://github.com/searxng/searxng), så søgeord ikke sendes direkte til en tredjepart. SearXNG skal have JSON-formatet slået til (`search.formats: [html, json]` i `settings.yml`), og hosten skal stå under Tilladte hosts. Organisationen kan låse søgetjenesten i maskinpolitikken.
8. **Audit-log** – metadata (host, model, tokens) uden indhold; giver mest mening med bridge. Slås til/fra i fanen, og **Vis log** viser de seneste poster (bridge: `audit.log` i konfigurationsmappen, maks. 1 MB med rotation til `.old`; direkte tilstand: de seneste 200 i add-in'et).
9. **Tilladte hosts** – politikken, én host pr. linje. Standard tillader kun loopback; `*.domæne`-wildcards kan bruges. I bridge-tilstand med maskinpolitik er feltet låst: *Politikken er fastsat af din organisation.*
10. **Mappe med brugeropgaver** (kun bridge) – mappen med opgaver fra [Brugerdefinerede opgaver](#brugerdefinerede-opgaver); tom = `<konfig>/tasks`.

Klik **Gem** – uden dette virker ændringerne ikke.

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

### Chat (alle tre apps)

Spørg om dit dokument eller få hjælp til Office. **Brug markering/dokument som kontekst** (til som udgangspunkt) giver modellen din aktuelle markering – i Excel dit markerede område – med. De seneste 20 beskeder medsendes som historik. Svar kan indsættes **efter** markeringen, **erstattes** eller indsættes **som kommentar**. **Ny samtale** nulstiller historikken. Enter sender, Shift+Enter linjeskift.

- **Figurer i chatten**: Hvis svaret indeholder et diagram (` ```mermaid `) eller en selvstændig SVG (```svg), vises det som billede i task pane, og knappen **Indsæt som billede** indsætter det – i Word som inline-billede ved markeringen, i PowerPoint som billedfyldt rektangel på den valgte slide (PowerPointApi 1.8).
- **Web-søgning**: Modellen kan selv søge online, når den mangler aktuelle oplysninger (indstillingen *Web-søgning*, fra som udgangspunkt). Undervejs vises "Søger på webbet: …", og svaret citerer sine kilder. Søgeresultater er utroværdig input og behandles som data – kritiske fakta bør verificeres.

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
| **Lokal app** (.dmg/.msi) | Kunden installerer MIAI App (Installation D ovenfor); bridgen kører lokalt, manifest sideloades automatisk | Ingen – alt lokalt, ingen internet |

Model-URL og valg af model sker i add-in'ets fanen **Indstillinger** i begge udgaver.

## Fejlfinding

| Symptom | Løsning |
|---|---|
| Gul badge / "CORS"-fejl i fejlbeskeden | Værktøjet tillader ikke kald fra add-in'ets adresse. Ollama: `OLLAMA_ORIGINS` (se Installation C). LM Studio: "Enable CORS". vLLM: `--allowed-origins`. Alternativt brug bridgen, som kører uden CORS. |
| "Endpoint er ikke tilladt af politikken." | Endpointet ligger uden for *Tilladte hosts*. Tilføj hosten (fx `llm.firma.local`) og **Gem** – medmindre en maskinpolitik låser feltet. |
| Rød badge "MIAI Bridge svarer ikke" | Bridgen kører ikke: `npm start` i bridge-tilstand, og brug bridge-manifestet (`packages/addin/manifest.xml`). |
| Ingen modeller i udfyldningsfeltet / "Test forbindelse" fejler | Kører værktøjet på den rigtige port? Prøv `curl http://127.0.0.1:11434/v1/models`. |
| Generering stopper uden svar | Tidsgrænsen er 120 s pr. endpoint. Vælg en mindre model eller en hurtigere hardware. |
| Formel giver `#NAME?` efter indsættelse | Modellen brugte et funktionsnavn, din Excel ikke kender. Ret formelen manuelt eller klik **Fortryd**. |
| Indstillinger genværrer ikke | Direkte tilstand gemmer i `localStorage`; nogle Mac-bygninger rydder den. Start på ny og gem igen, eller brug bridgen (gemmer i `config.json`). |
| Task pane åbner ikke | Genstart Office efter sideload. Fjern evt. add-in'et (Indsæt → Tilføjelser → …) og sideload igen. |

## Se også

- [Udvikling og afprøvning](UDVIKLING.md) – byg, tests, konfigurationsfiler
- [Arkitektur](ARKITEKTUR.md) – komponenter, dataflow, sikkerhed
- [Projektplan](PLAN.md) – faser og funktionskatalog
- [Arkitekturbeslutninger](adr/) – bl.a. ADR-0006 (politik) og ADR-0007 (bridge vs. direkte)
