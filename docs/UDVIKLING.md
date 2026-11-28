# Udvikling og afprøvning

Vil du bare bruge MIAI, så følg [installationen i brugervejledningen](BRUG.md#installation--vælg-den-vej-der-passer-til-dig). Denne side er til dig, der vil bygge, ændre eller udgive.

## Forudsætninger
- Node.js 22 LTS
- Microsoft 365 / Office 2021+ desktop (Windows eller Mac)
- Et OpenAI-kompatibelt inference-værktøj, fx [Ollama](https://ollama.com):
  ```bash
  ollama pull qwen3:4b
  ```
  Alternativer: LM Studio (start "Local Server"), `llama-server -m model.gguf --port 8080`, vLLM, LocalAI, Jan m.fl.

## Hostet add-in på GitHub Pages
Når GitHub Pages er slået til (Settings → Pages → Source: **GitHub Actions**), publicerer `pages.yml` add-in'et ved hvert push til `main` sammen med færdige manifester:

- `https://<bruger>.github.io/<repo>/manifest.xml` (Word, Excel, PowerPoint)
- `https://<bruger>.github.io/<repo>/manifest-outlook.xml` (Outlook)

I dette repo er adressen `https://mor10w.github.io/MIAIPlugin/`. Sådan tages manifestet i brug, står i [brugervejledningen, vej B](BRUG.md#b-hostet-add-in-uden-app). I en fork skal inference-værktøjet tillade CORS fra forkens egen adresse.

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
Bridgen kører på `https://localhost:44321` og serverer også task pane. Er porten optaget, venter bridgen op til 5 sekunder på den (typisk er det den forrige udgave, der er ved at lukke). Er den stadig optaget (fx af en anden brugers bridge), vælges og huskes den næste ledige, og MIAI App opdaterer manifestet i Office. `MIAI_PORT` låser porten.

API'et kræver et adgangstoken pr. bruger (#11). Bridgen skriver dit personlige manifest med port og token i konfigurationsmappen og viser stien ved start. Sideload det:
```bash
npm run sideload:mac -- ~/Library/Application\ Support/MIAI/manifest.xml
```
På Windows: `npm run sideload:windows -- %APPDATA%\MIAI\manifest.xml`. Linux: `~/.config/miai/manifest.xml`.

Watch-tilstand for både add-in og bridge: `npm run dev:bridge`.

### Regnearksfunktionerne med bridgen fra kildekoden

`=MIAI.ASK()` og `=MIAI.MODELS()` er med i alle tre manifester til Word, Excel og PowerPoint. Udviklingsmanifestet kører direkte mod modellen og kræver derfor, at modelserveren tillader CORS fra `https://localhost:3000`. Med en server uden CORS (fx oMLX), eller for at prøve websøgning og egne mapper i cellen, bruges bridgen fra kildekoden. Den bruger samme konfigurationsmappe som MIAI App, så endpoints, modeller og adgangstoken er de samme, og udviklingscertifikatet fra `npm run certs` (det samme som `npm run dev`).

1. Afslut MIAI App (menuen → Afslut), så porten er fri. Kører appen stadig, vælger bridgen en anden port, og det manifest, Office allerede har, peger så på en port, hvor intet lytter.
2. Byg og start bridgen:
   ```bash
   npm run build
   ```
   ```bash
   npm start
   ```
3. Læg det personlige manifest, bridgen har skrevet, ind i Office, og genstart Excel:
   ```bash
   npm run sideload:mac -- ~/Library/Application\ Support/MIAI/manifest.xml
   ```
   Windows: `npm run sideload:windows -- %APPDATA%\MIAI\manifest.xml`.

   Gør det igen, hver gang bridgen melder en anden port ved start. Manifestet i Office følger ikke selv med.
4. Har Excel på denne maskine før brugt funktionerne fra et andet MIAI-manifest (typisk udviklingsmanifestet), så ryd Excels funktionscache, mens Excel er lukket. Se [Skift mellem manifester med funktionerne](#skift-mellem-manifester-med-funktionerne).
5. Under Indstillinger vises nu **Excel-funktioner**, og `=MIAI.ASK(` findes i arket. Det første kald kan tage op mod et minut, mens modellen indlæses.

**Tilbage til MIAI App:** stop bridgen med Ctrl+C, og start MIAI App. Den skriver manifestet igen og lægger det ind i Office. Genstart Excel.

### Skift mellem manifester med funktionerne

Udviklingsmanifestet, bridge-manifestet og det hostede manifest har hver sit id, men samme navnerum (`MIAI`). Excel gemmer, hvilket add-in navnerummet hører til, og laver ikke registreringen om, når et andet manifest med samme navnerum lægges ind. Den gamle registrering peger så på et add-in, der ikke længere er der.

Sådan ser det ud:

- Panelet viser **Excel-funktioner** under Indstillinger, men `=MIAI.ASK(` foreslås ikke i arket, eller cellen viser `#NAME?`.
- Eller cellen står på `#BUSY!`, og hverken bridgen eller modelserveren får et kald. Tidsgrænsen på 20 sekunder slår ikke til, fordi Excel aldrig kalder funktionen.

Ryd funktionscachen, hver gang du skifter mellem `npm run dev`, bridgen og det hostede add-in:

1. Luk Excel helt (Cmd+Q på Mac).
2. Slet mappen med Excels registreringer af regnearksfunktioner. Det er ren cache, som Excel bygger igen ved næste start.

   macOS:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef/CustomFunctions
   ```
   Windows: slet indholdet af `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (Microsofts anvisning til at rydde Office-cachen; ikke afprøvet her).
3. Start Excel, åbn MIAI-panelet, og prøv formlen igen.

Tjek på macOS, hvilket add-in der har registreringen:

```bash
cat ~/Library/Containers/com.microsoft.Excel/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef/CustomFunctions/*/Registry/*/*/properties.json
```

`addinid` skal indeholde id'et fra det manifest, du har lagt ind (`fdac46a3-…` for bridgen, `3b0f5a2e-…` for udviklingsmanifestet, `9c7e2d14-…` for det hostede).

## Tests
```bash
npm test
```
- `shared`: opgaver, tekst-hjælpere og motoren i direkte tilstand mod en mock-server.
- `bridge`: streaming, `<think>`-filter, JSON-validering med genforsøg, politik, sikkerhedstjek, brugeropgaver og audit-log.
- `addin`: hjælpefunktioner (SVG, billedstørrelser, Excel-oversigt).

**Udgivelse af MIAI App (ADR-0016):**
Hvert push til `main`, der rører app-koden, udgiver automatisk en ny release (`auto-release.yml`): patch-versionen bumpes (fx 0.1.2 → 0.1.3), committes til `main` og tagges, så klienter med MIAI App henter opdateringen i baggrunden. Vil du udgive en bestemt version (fx 0.2.0), ret `version` i `packages/app/package.json` og commit: push'et udgiver den præcist uden bump. Bagefter sker det som før:
1. `release.yml` bygger macOS (dmg og zip, arm64 og x64) og Windows (NSIS, x64), signerer, hvis secrets er sat op, og laver en GitHub Release.
2. Pages kopierer filerne og opdateringsfeedet (`latest*.yml`) til `<pages-url>/updates/`, hvor appen henter opdateringer.

**Udviklerudgave (ADR-0016, tillæg):** hvert push til `main`, der rører appen, bygger en `.dmg` til Apple Silicon og en installer til Windows (x64) i `dev-build.yml` med version `<version>-dev.<kørselsnummer>` og lægger dem som forhåndsudgivelsen `dev`. Der udgives kun, når begge er bygget. Pages serverer dem under `<pages-url>/dev/`, og appens menu viser *Hent udviklerudgave …*, når der er en nyere til din platform. Workflowet kan også startes manuelt (Actions → Udviklerudgave); fra andre grene end `main` bygges der kun, og filerne ligger som artefakter på kørslen. Tjekliste på en rigtig Mac:
- Menupunktet dukker op i en ældre udgave og åbner hentningen i browseren.
- Appen kan godkendes i Systemindstillinger → Anonymitet & sikkerhed → *Åbn alligevel* og starter bagefter.
- Den nye udgave viser *Udviklerudgave X er installeret*, og nøgler i nøgleringen og Office-installationen virker stadig.

På en rigtig Windows-pc:
- Menupunktet dukker op i en almindelig udgave og henter `MIAI-Setup-…-dev.N-x64.exe` i browseren.
- SmartScreen kan omgås med *Flere oplysninger* → *Kør alligevel* (indtil #67), og installeren installerer oven i den gamle udgave uden admin.
- Den nye udgave viser *Udviklerudgave X er installeret*, og certifikat, nøgler og Office-installationen virker stadig.
- Efter næste push til `main`: *Søg efter opdateringer…* henter den nyere udviklerudgave fra `dev.yml`, beskeden *MIAI er klar til at opdatere* vises, og et klik installerer den og starter MIAI igen.

Lokalt: kør først `npm run build` (appen pakker add-in'et og bridgen med), og derefter `npm run dist:mac -w @miai/app` eller `npm run dist:win -w @miai/app`. Filerne ender i `packages/app/dist/`. Mac-bygget er ad hoc-signeret som udgivelserne; en helt usigneret app starter ofte ikke på Apple Silicon. Windows-bygget er usigneret. Et lokalt byg er ikke hentet fra nettet, så Gatekeeper blander sig ikke; det gør det nyttigt, når en hentet udgave ikke vil starte. Release-workflowet kan også startes manuelt (Actions → Release); så bygges der kun, uden release.

### Licenser i de byggede filer
MIAI er udgivet under MIT (`LICENSE`), og de pakker, den bygger på, kræver, at deres licenstekst følger med, når koden gives videre. Kildekoden henviser kun til pakkerne; det er de byggede filer, der indeholder dem. Derfor laves licensfilerne ved hvert byg og ligger ikke i repoet:

- **Add-in'et:** `npm run build` lægger `LICENSE.txt` og `THIRD-PARTY-NOTICES.txt` i `packages/addin/dist`. Listen laves af et Vite-plugin ud fra de moduler, der faktisk er pakket med, og følger derfor afhængighederne af sig selv. Filerne serveres sammen med add-in'et, fx `<pages-url>/THIRD-PARTY-NOTICES.txt`.
- **MIAI App:** `npm run notices -w @miai/app` skriver `packages/app/build/THIRD-PARTY-NOTICES.txt` med alt, appen afhænger af ved kørsel, samt Electron. electron-builder lægger den og `LICENSE.txt` i appens `Resources` ved siden af `addin/`. Chromiums og Electrons egne licenser lægger electron-builder selv ved. `dist:mac`, `dist:win` og workflowene kører trinnet; udviklerbygget kontrollerer, at filerne er med.

Begge dele bruger `scripts/third-party-notices.mjs`. En pakke uden licensfil nævnes med den licens, den selv angiver, og en bemærkning om, at teksten mangler.

**CI:** `ci.yml` bygger og tester på Linux og Windows og kører lint og e2e. Det er de tjek, der skal være grønne. macOS bygges i `ci-macos.yml` som et separat, ikke-blokerende tjek: GitHubs Mac-runnere er ofte optaget, og et job, der aldrig får en maskine, skal ikke gøre CI rødt. Fejler macOS-tjekket på selve koden, skal det rettes som alt andet. Står det i kø eller bliver annulleret, så kør det igen fra Actions.

### Browsertests (task pane)
```bash
npm run build
```
```bash
npm run test:e2e
```
Kører add-in'ets build i Chromium med Playwright (`packages/addin/e2e/`). Testene kører også i CI (jobbet `e2e`), hvor skærmbilleder og traces gemmes som artefakter ved fejl.

- **Simuleret Office.js** (`e2e/fixtures/office.ts`): Word, Excel og PowerPoint med de API'er, adapterne bruger. Kalder en adapter et API, simuleringen ikke kender, fejler det med "Simuleret Office.js mangler: …" – udvid så simuleringen. Alt, adapterne gør ved dokumentet, registreres og kan efterprøves (`tp.calls()`, `tp.find("picture")`). Den simulerede Excel tæller også trinene i Excels egen fortryd (`window.__miai.undoGroups()`): ét pr. `context.sync()` med ændringer, eller ét pr. `Excel.run` med `mergeUndoGroup` (kræver `apiSets: { ExcelApi: "1.20" }`). Formelsporing (#76): simuleringen regner ikke formler ud, så formlernes resultater gives med arket (`results: { B5: "#DIV/0!" }`), og præcedenter og afhængige udledes af formlernes cellehenvisninger. Som i Excel fejler et kald uden præcedenter først ved `context.sync()` (`ItemNotFound`), og så er intet fra samme rundtur indlæst. Kaldene registreres som `{ kind: "trace", method, address }`.
- **Mock-model** (`e2e/fixtures/servers.ts`): OpenAI-kompatibel server med streaming, ikke-streaming og værktøjskald, med eller uden CORS. Alle forespørgsler gemmes (`tp.model.lastRequest`).
- **Tilstande:** direkte (add-in hostet statisk over HTTPS) og bridge (den rigtige, byggede bridge med midlertidig konfiguration). Alle porte er dynamiske, så en lokal Ollama ikke forstyrrer.
- Testene fejler også, hvis task pane kaster en uventet JavaScript-fejl.

Lokalt bruges den Chromium, Playwright har installeret (`npx -w @miai/addin playwright install chromium`).

Browsertestene erstatter ikke den manuelle test i rigtig Office (testlisten nedenfor). De fanger regressioner i UI og adaptere, før den manuelle runde.

### Demovideo
`packages/addin/e2e/demo/` optager en kort demovideo af task pane ved siden af et simuleret dokument: korrektur som sporede ændringer, agenten med plan og godkendelse, og chat i Word, derefter agenten i Excel (tomme rækker, tabelformat, datalinjer og diagram) og til sidst design i PowerPoint (master og layouts). Den bruger den simulerede Word, Excel og PowerPoint fra browsertestene og indspillede modelsvar, så optagelsen er ens fra gang til gang og kan laves om, når brugerfladen ændrer sig. Den viser altså brugerfladen, ikke en models kvalitet eller hastighed, og scenen siger det selv i øverste hjørne.

```bash
npm run build
```
```bash
npx playwright test -c packages/addin/e2e/demo/playwright.config.ts
```
Optagelsen (1920×1080, knap to minutter) ender i `packages/addin/demo-results/`. Til MP4:
```bash
ffmpeg -i packages/addin/demo-results/demo.video.ts-demovideo/video.webm -c:v libx264 -crf 20 -pix_fmt yuv420p -movflags +faststart -an miai-demo.mp4
```
Dokumentet, regnearket, præsentationen, modelsvarene og rækkefølgen står øverst i `demo.video.ts`; scenen (dokument, regneark, slides, billedtekster, markør) er `stage.html`. Slides fra designet vises med MIAI's egen forhåndsvisning, fordi simuleringen ikke kan tegne en .pptx; den bygger på de samme skabeloner som den master, MIAI sætter ind. `MIAI_DEMO_PROBE=1` kører det hele igennem på få sekunder og gemmer et stillbillede pr. scene.

**Videoen om `=MIAI.ASK()`** (`ask.video.ts`, knap halvandet minut) viser regnearksfunktionen til brug i opslag: formlen skrives i en celle, trækkes ned over rækkerne, en ny prompt giver en ny kolonne, svaret følger med, når teksten rettes, og til sidst slår en celle op på nettet med `FALSK` som sidste argument. Regnearket er scenens eget (`stage.grid` i `stage.html`): scenen spiller Excels rolle, viser `#BUSY!` og kalder den funktion, add-in'et har registreret i task pane. Task pane serveres af den rigtige bridge, fordi websøgning i en celle kun findes dér. Køen, cachen, statuslinjen "MIAI.ASK: … i gang, … i kø" og søgerunden er derfor MIAI's egen kode; kun gitteret, modelsvarene og søgeresultaterne er simulerede, og optagelsen rører ikke internettet. Henvendelserne, virksomhederne, promptene og svarene står øverst i filen.

```bash
npx playwright test -c packages/addin/e2e/demo/playwright.config.ts ask.video.ts
```
```bash
ffmpeg -i "packages/addin/demo-results/ask.video.ts-video-MIAI-ASK-/video.webm" -c:v libx264 -crf 20 -preset slow -pix_fmt yuv420p -movflags +faststart -an docs/demo/miai-demo-excel-ask.mp4
```
Uden filnavn optager kommandoen begge videoer. Den færdige video ligger i [`docs/demo/miai-demo-excel-ask.mp4`](demo/miai-demo-excel-ask.mp4).

Videoen beviser ikke, at det samme virker i rigtig Office med en rigtig model. [`docs/demo/TJEKLISTE.md`](demo/TJEKLISTE.md) gentager videoens forløb i rigtig Word og Excel med testfilerne i samme mappe; kør den, før videoen bruges udadtil.

## Manifester
```bash
npm run manifests
```
Regenererer `manifest.xml` (bridge) og `manifest.dev.xml` fra `manifest.template.xml`, og `manifest.outlook.xml` og `manifest.outlook.dev.xml` fra `manifest.outlook.template.xml` (Outlook, MailApp). Hostede varianter:
```bash
node scripts/make-manifest.mjs hosted https://intern.server/miai/ manifest.xml
```
```bash
node scripts/make-manifest.mjs hosted https://intern.server/miai/ manifest-outlook.xml outlook
```
**Regnearksfunktioner (#80, ADR-0018):** alle tre manifester til Word, Excel og PowerPoint (bridge, dev og hostet) har Excels shared runtime og funktionerne `=MIAI.ASK()` og `=MIAI.MODELS()` (`functions: true` i `scripts/make-manifest.mjs`). Outlook-manifesterne har dem ikke. Bridge-manifestet er afprøvet i rigtig Excel på Mac; Windows, Excel på web og det hostede manifest er ikke afprøvet endnu. Et manifest med en skemafejl indlæses slet ikke, heller ikke i Word og PowerPoint, så sæt `functions: false` på en variant og kør `npm run manifests`, hvis den skal udgives uden funktionerne. Et manifest med funktionerne åbner task pane med `?functions=1`; uden den parameter registrerer siden ingen funktioner og viser intet om dem. Metadata ligger i `packages/addin/public/functions.json` (dansk) og `functions.en.json`. Office cacher funktionernes metadata; ryd Office-cachen, hvis en ændring ikke slår igennem. Excel cacher også, hvilket add-in navnerummet `MIAI` hører til; se [Skift mellem manifester med funktionerne](#skift-mellem-manifester-med-funktionerne).

Outlook sideloades fra fil: Hent tilføjelsesprogrammer → Mine tilføjelsesprogrammer → Tilføj brugerdefineret → Tilføj fra fil. Brug `manifest.outlook.dev.xml` med `npm run dev`.

**Central konfiguration (direkte tilstand):** en `miai-config.json` ved siden af `taskpane.html` giver organisationens låste politik og standardindstillinger (ADR-0015). I dev kan filen lægges i `packages/addin/public/`.

## Konfiguration
**Direkte tilstand**: indstillinger gemmes i add-in'ets `localStorage` på maskinen, og de redigeres under fanen Indstillinger.

**Bridge**:

| | Sti |
|---|---|
| Brugerindstillinger | Windows `%APPDATA%\MIAI\config.json` · Mac `~/Library/Application Support/MIAI/config.json` · Linux `~/.config/miai/config.json` |
| Maskinpolitik (organisation) | Windows `%ProgramData%\MIAI\policy.json` · Mac `/Library/Application Support/MIAI/policy.json` · Linux `/etc/miai/policy.json` |
| Audit-log (hvis slået til) | `audit.log` ved siden af `config.json` |
| Vidensindeks (F15) | `knowledge.db` ved siden af `config.json` (SQLite via `node:sqlite`) |
| API-nøgler | OS'ets nøglering, tjeneste `MIAI` (konto `endpoint:<id>`). `MIAI_KEYRING=off` gemmer dem i `config.json` i stedet; unit-testene sætter den, så udviklerens nøglering ikke røres. |

Miljøvariabler for bridgen: `MIAI_PORT` (44321), `MIAI_CONFIG_DIR`, `MIAI_POLICY_FILE`, `MIAI_TLS_CERT`, `MIAI_TLS_KEY`, `MIAI_ADDIN_DIR`, `MIAI_LOG`.

Eksempel på maskinpolitik, der kun tillader en intern inference-server:
```json
{ "allowedHosts": ["127.0.0.1", "localhost", "llm.firma.local"], "locked": true, "webSearch": false }
```
Eller med søgning via organisationens egen SearXNG:
```json
{ "allowedHosts": ["127.0.0.1", "localhost", "llm.firma.local", "search.firma.local"], "locked": true, "search": { "provider": "searxng", "url": "https://search.firma.local" } }
```

## Fase 0 – manuel testliste
Skal afprøves i Office desktop. Trin for trin, med testdata og resultatskabeloner: [TESTVEJLEDNING.md](TESTVEJLEDNING.md).

**Forbindelse**
- [ ] Direkte, dev (`https://localhost:3000`) → Ollama uden CORS-opsætning, i Windows (WebView2) og Mac (WKWebView)
- [ ] Direkte, hostet (GitHub Pages) → Ollama med `OLLAMA_ORIGINS`: kommer der en Local Network Access-prompt i WebView2?
- [ ] Direkte → LM Studio med og uden "Enable CORS": fejlteksten skal forklare opsætningen
- [ ] Bridge-tilstand detekteres automatisk; badge viser "bridge"
- [ ] `localStorage` bevares mellem genstarter af Word/Excel/PowerPoint (især på Mac)

**Office på nettet: lokal netværksadgang (#74)**

Afprøves i nyeste Chrome og Edge med det hostede add-in (GitHub Pages) mod Ollama med `OLLAMA_ORIGINS`. Koden bygger på specifikationen og er ikke afprøvet i en rigtig browser; notér afvigelser, og ret teksterne `lnaDenied`, `lnaFrame` og `lnaPrompt` (`i18n.ts`) og afsnittet i [BRUG.md](BRUG.md#office-på-nettet-lokal-netværksadgang-i-chrome-og-edge), hvis browserens ordlyd er en anden.

- [ ] Første kald: spørger browseren om lokal netværksadgang, også uden et klik fra brugeren? Notér ordlyden, og hvilken side tilladelsen gemmes for (Office-siden eller add-in'ets adresse)
- [ ] `await navigator.permissions.query({ name: "loopback-network" })` i konsollen, i add-in'ets ramme: svarer `prompt`, `granted` eller `denied` som forventet? `document.featurePolicy.allowsFeature("loopback-network")`: giver Office på nettet tilladelsen videre til add-in'et?
- [ ] **Tillad**: status bliver *Forbundet* inden for 15 sekunder (eller efter **Søg igen**), og en opgave kører
- [ ] **Bloker**: status *Blokeret af browseren*, banner med vejledning, guiden viser blokeringen én gang, og en opgave fejler med samme overskrift (ikke "Kunne ikke forbinde" med CORS-vejledning)
- [ ] Fortryd via webstedsindstillinger som beskrevet i BRUG.md, genindlæs: forbindelsen virker
- [ ] Spørgsmålet lukkes eller ignoreres: guiden viser "Intet værktøj fundet" og hjælpeteksten om at vælge Tillad. Kommer spørgsmålet igen ved hvert statustjek (hvert 15. sekund), og forsvinder det, når et kald afbrydes af timeout?
- [ ] Giver Office på nettet ikke tilladelsen videre: MIAI viser beskeden om, at siden ikke giver add-in'et lov (ikke vejledningen til webstedsindstillinger)
- [ ] Politikken `LocalNetworkAccessAllowedForUrls` med add-in'ets adresse og derefter med Office-sidens adresse: hvilken af dem fjerner spørgsmålet?
- [ ] MIAI App (bridge) brugt i Office på nettet: spørger browseren, før task pane hentes fra `https://localhost`, og hvad sker der ved Bloker?

**Funktioner**
- [ ] Word: ret grammatik → sporede ændringer vises og kan accepteres/afvises
- [ ] Word: omskriv, oversæt, opsummér (uden markering = hele dokumentet), artikel indsættes som formateret tekst, kommentar
- [ ] Excel: forklar formel; generér formel → validering og **Fortryd**; oversæt område → celler; kommentar
- [ ] PowerPoint: omskriv tekst i markeret tekstboks; tekst til slides med standardskabelon **og** en firmaskabelon
- [ ] Chat i alle tre apps; stop-knappen afbryder generering
- [ ] Vedhæft et skærmbillede (klips og Ctrl+V) med en vision-model: svaret beskriver billedet; "indsæt tabellen fra billedet i arket" giver en plan med de rigtige tal; med en model uden vision vises beskeden om at vælge en anden model
- [ ] Vedhæft en PDF og en .docx med bridge: indholdet bruges i svaret; uden bridge forklares, at det kræver MIAI Bridge
- [ ] Markeringen under chattens skrivefelt følger med, når der markeres noget andet (Word, Excel, PowerPoint; Windows og Mac)
- [ ] Samtalen er der igen, når task pane lukkes og åbnes på samme gemte dokument (også OneDrive/SharePoint); "Ny samtale" sletter den; et nyt, ikke gemt dokument huskes ikke
- [ ] Endpoint uden for politik afvises; låst maskinpolitik (bridge) kan ikke ændres
- [ ] Content-Security-Policy (#28): add-in'et virker som før i Word, Excel, PowerPoint og Outlook på Windows, Mac og web (Office.js indlæses, opgaver kører, figurer vises og indsættes); konsollen viser kun blokerede kald til Microsofts telemetri, ikke til modellen
- [ ] Tilføj en host under Tilladte hosts og gem: task pane genindlæses, Indstillinger åbner igen, og den nye host kan kaldes
- [ ] Mørkt Office-tema
- [ ] Reasoning-model (Qwen3): ingen `<think>` i output
- [ ] Kontekst (#73): med Ollama uden `OLLAMA_CONTEXT_LENGTH` viser Indstillinger → Modeller den indlæste kontekst (4.096 på en maskine med under 24 GB VRAM) og en advarsel; efter `OLLAMA_CONTEXT_LENGTH=32768` og genstart viser **Tjek igen** 32.768 og ingen advarsel. I LM Studio vises den kontekst, modellen er indlæst med
- [ ] Regnearksfunktioner (#80), i Excel på Windows og Mac, med både MIAI App (bridge) og `npm run dev` eller det hostede add-in (direkte):
  - [ ] Manifestet indlæses, og MIAI virker som før i Word, Excel og PowerPoint (task pane åbner fra fanen Hjem)
  - [ ] Har maskinen før haft et andet MIAI-manifest med funktionerne, er Excels funktionscache ryddet først (se [Skift mellem manifester med funktionerne](#skift-mellem-manifester-med-funktionerne))
  - [ ] `=MIAI.ASK("Hvilket land ligger byen i?"; A2)` viser `#BUSY!` og derefter svaret; formlen kan trækkes ned over 50 rækker, og task pane viser "i gang" og "i kø" med et virkende **Stop MIAI-funktioner**
  - [ ] Mens cellerne regner, viser panelet statuslinjen "MIAI.ASK: … i gang, … i kø" (gør det ikke, kører Excel funktionerne i en anden udgave af siden end panelet; notér det)
  - [ ] Mens formlen skrives, kaldes modellen ikke (ingen aktivitet i Ollama, før Enter)
  - [ ] F9 og Ctrl+Alt+F9 spørger ikke modellen igen for uændrede celler
  - [ ] Gem, luk og åbn projektmappen: kaldes funktionerne igen? Notér resultatet her: ______
  - [ ] En `MIAI.ASK`-celle som input til en anden giver ét modelkald pr. celle
  - [ ] `optimized = SAND` er mærkbart hurtigere; uden det og med websøgning slået til (bridge) søger cellen på nettet
  - [ ] **Indsæt modelvælger** laver en dropdown i den markerede celle med modellerne (accepterer Excel kilden `=MIAI_modeller!$A$1#`?), og `=MIAI.ASK("…"; A2; ; $H$1)` bruger den valgte model
  - [ ] Et ukendt modelnavn giver `#VALUE!` med forklaringen i fejlmenuen
  - [ ] Stop Ollama: cellen viser `#N/A` med "MIAI kan ikke nå modellen"
  - [ ] Windows: kald til Ollama på localhost virker fra funktionen (loopback, #9)

**Outlook og central udrulning (#7)**
- [ ] Outlook (Windows, Mac, nyt Outlook, web): manifestet tilføjes fra fil; knappen AI-assistent vises i en læst mail og i et udkast
- [ ] Læst mail: Svar på mail åbner et udkast med formateret svar (fed, punkter); Svar alle virker; intet sendes automatisk
- [ ] Udkast: Omskriv med markering erstatter kun markeringen, med formatering; uden markering bruges hele udkastet
- [ ] Udkast i ren tekst (ikke HTML): indsættes som ren tekst uden HTML-koder
- [ ] Central udrulning: add-in hostet internt, `miai-config.json` ved siden af, manifest uploadet i M365 Admin Center → add-in'et vises hos brugeren (også Office på web), politikken står som "Organisationen" og kan ikke ændres, og standard-endpointen er den interne server
- [ ] Office på web: kald til den interne inference-server virker (CORS fra add-in'ets adresse)
- [ ] Ugyldig `miai-config.json`: politikken låses, intet uden for standardhosts tillades

**MIAI App (Windows)** (installer fra Release-workflowet):
- [ ] Installation uden admin; MIAI starter og ligger i proceslinjen med farvet ikon
- [ ] Første start: Windows' dialog om certifikatet vises, og efter "Ja" står det under certmgr → Betroede rodnøglecentre (Aktuel bruger)
- [ ] Office: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` har MIAI's Id og stien til det personlige manifest; efter genstart findes AI-assistent på fanen Hjem i Word, Excel og PowerPoint
- [ ] "Start automatisk ved login" virker efter log ud/ind
- [ ] Outlook: "Vis Outlook-manifest i Stifinder" åbner mappen; filen kan tilføjes i Outlook
- [ ] Opdatering: installér en ældre version, udgiv en nyere → beskeden "MIAI er klar til at opdatere" vises (efter *Søg efter opdateringer…* eller af sig selv), menuen viser "Genstart og opdater til …", og den nye version kører bagefter
- [ ] *Søg efter opdateringer…* i den nyeste version: beskeden "Du har den nyeste version"
- [ ] Signeret build (når Azure Trusted Signing er sat op): ingen SmartScreen-advarsel; udgiver vises korrekt

**MIAI App (Mac)**
- [ ] "Start automatisk ved login" til: MIAI står under Systemindstillinger → Generelt → Login-emner og starter efter log ud/ind (én instans, ingen dubletter)
- [ ] Slå det fra igen: appen bliver kørende og starter ikke ved næste login
- [ ] Startet fra DMG'en (ikke i Programmer): appen tilbyder at flytte sig til Programmer
- [ ] Opgradering fra en udgave med LaunchAgent (`~/Library/LaunchAgents/dk.miai.app.plist`): filen fjernes, login-elementet er slået til, og appen starter kun én gang ved login
- [ ] Ny udgivelse: beskeden "Ny version af MIAI" vises, menuen viser *Hent MIAI X…*, og klikket henter den rigtige dmg (arm64/x64) fra `<pages-url>/updates/`
- [ ] Installér i Office uden adgang: dialogen om Fuld diskadgang vises, og "Åbn Indstillinger" åbner det rigtige sted

**Agent i Word – formatering** (simuleringen efterligner Words regler, men kun rigtig Word kan bekræfte dem). Brug et dokument med titel, overskrifter, fed tekst, et link, en punktliste, en nummereret liste og en tabel:
- [ ] Erstat et ord midt i en sætning med fed tekst: kun ordet skifter, og fed tekst omkring det bevares
- [ ] Indsæt efter en overskrift: det nye afsnit er brødtekst (ikke overskrift), uden arvet fed eller farve; fed, kursiv og link i teksten virker
- [ ] Indsæt "- punkt" efter et almindeligt afsnit: bliver et punkt i dokumentets punktstil; "1. …" bliver nummereret
- [ ] Indsæt efter det sidste punkt i en liste med typografien brødtekst: afsnittet er ikke med i listen
- [ ] Skift typografi (brødtekst → overskrift 2 og punkt → brødtekst): teksten og dens fed/kursiv er urørt
- [ ] Slet et helt afsnit: intet tomt afsnit bliver tilbage; slet et punkt: intet tomt punkt
- [ ] Flyt et afsnit med fed tekst og et link: formateringen følger med, og der opstår intet ekstra tomt afsnit
- [ ] Tabeller: tekst i en celle kan rettes; sletning/flytning af celleafsnit er fravalgt i planen
- [ ] Tabeller (#24): "tilføj en række" og "slet rækken med …" ændrer den rigtige række (også i en tabel med flettede celler); "lav en tabel med …" giver en Word-tabel med overskriftsrække og dokumentets tabeltypografi; sporede ændringer og Fortryd virker for alle tre
- [ ] En Markdown-tabel i et chatsvar bliver en rigtig tabel ved "Indsæt efter"
- [ ] Fortryd med sporede ændringer (Microsoft 365): kun agentens ændringer og kommentarer forsvinder; egne sporede ændringer før sessionen er urørte
- [ ] Fortryd i Word 2016/2019 (uden WordApi 1.6): dokumentet er som før; nummerering i lister og sidehoveder er intakte
- [ ] Ændr et afsnit mellem plan og "Udfør": handlinger på det afsnit afvises med "Dokumentet er ændret siden planen"

**Agent i Excel – formatering** (ark med en Excel-tabel, valuta- og procentformat, fed overskrift og fyldfarve):
- [ ] "Tilføj en række": rækken lægges i tabellen (tabellen udvides), med tabellens formatering og beregnede kolonner
- [ ] "Tilføj en total under tabellen": totalen får kolonnens talformat (fx kr.), ikke "Standard"
- [ ] Skriv i celler, der allerede har formatering: formateringen er urørt
- [ ] Fortryd: værdier, talformater, fed og fyld er som før, og tabellen har sin oprindelige størrelse
- [ ] "Fjern de tomme rækker i tabellen": tabellen krymper, data under den rykker op; Fortryd sætter rækkerne ind igen med værdier og formater
- [ ] Tomme rækker i et område uden tabel: kun områdets kolonner rykker op, data til højre står stille
- [ ] "Gør tabellen pæn med et moderne layout": området bliver en Excel-tabel med en indbygget typografi (eller en eksisterende tabel skifter typografi), talformat og kolonnebredde passer; Fortryd gør det til et almindeligt område med de oprindelige formater og kolonnebredder
- [ ] Samme ønske i **Chat**: svaret henviser til agenten (ingen VBA), og knappen "Gør det med agenten" vises
- [ ] Sortér en tabel og et område efter en kolonne (overskrift og bogstav): rækkerne følger med inkl. formater; Fortryd giver den oprindelige rækkefølge
- [ ] Betinget formatering (farveskala, datalinjer, større end, dubletter, top 10), diagram ved siden af data og frys overskrift; Fortryd fjerner dem og sætter frysningen tilbage
- [ ] Kodetilstand slået fra: et script fra modellen er fravalgt med "Kodetilstanden er slået fra"
- [ ] Pivottabel som fast værktøj (kodetilstand slået fra): "Lav en pivottabel med omsætning pr. region på et nyt ark" giver en rigtig pivottabel med de rigtige felter; Fortryd fjerner pivottabellen og arket
- [ ] Datavalidering: rulleliste og tal mellem to grænser virker i cellerne; Fortryd giver den tidligere regel (eller ingen) tilbage
- [ ] Kodetilstand slået til: "Opret et navngivet område for C2:C5": koden kan ses, scriptet kører, og Fortryd gendanner arket
- [ ] Et script, der fejler midtvejs: arket er som før, og "Bed modellen rette scriptet" giver et nyt script med fejlen rettet
- [ ] Flere ark: "Lav et nyt ark med den samlede omsætning" opretter arket med en formel på tværs af ark; skrivning på et andet ark end det aktive; omdøb et ark; Fortryd sletter det nye ark og sætter navn og celler tilbage (også med mellemrum i arknavnet)
- [ ] Flere trin: bed om noget, der kræver data fra et andet ark ("tilføj budgettet fra Budget-arket"): "Læser Budget!…" vises, og planen bruger de rigtige tal; intet er ændret, før planen udføres
- [ ] En handling, der fejler (fx sortering efter en kolonne, der ikke findes): "Bed modellen rette det, der fejlede" giver en ny plan kun for det fejlede
- [ ] Klikbare henvisninger: en celleadresse i et chatsvar markerer cellerne (også på et andet ark); "Vis i dokumentet" i planen markerer målet i Excel, Word og PowerPoint uden at ændre noget
- [ ] Henvisninger til afsnit i Word (#77): spørg i **Chat** om et længere dokument uden markering (med overskrifter, en tabel og tomme afsnit): svaret henviser med `[¶12]`, og et klik markerer det rigtige afsnit (også i en tabel); slet et afsnit længere oppe, og klik igen: det samme afsnit markeres
- [ ] Kodetilstand i Word: "Skriv 'Fortroligt' i sidehovedet" kører; et script, der fejler midtvejs, efterlader brødteksten som før; Fortryd gendanner brødteksten med formatering, billeder og tabeller
- [ ] Kodetilstand i PowerPoint (Microsoft 365): et script, der tilføjer en figur eller sletter slides, kan fortrydes, og alle slides er tilbage med billeder, noter og animationer; i PowerPoint 2019 afvises scriptet med en forklaring
- [ ] Politik `agentScripts: false` (bridge-politik eller `miai-config.json`): kontakten er slået fra og kan ikke ændres

**Agent i Excel – overskrivning, markering og ét fortryd (#75)** (kan kun bekræftes i rigtig Excel; simuleringen tæller kun trin i fortryd-stakken):
- [ ] Advarsel: en plan, der skriver oven i data og formler, viser "Overskriver N celler, heraf M formler" (og "Rydder …" ved ryd) med de rigtige tal; en plan, der kun skriver i tomme celler, viser ingen advarsel; arket er urørt, og Ctrl+Z har intet nyt trin, før der trykkes Udfør
- [ ] Ét fortryd (Microsoft 365, Windows 2509 / Mac 16.100 eller senere): efter en plan med flere handlinger (værdier, formel, ny tabelrække, nyt ark, omdøbning) fjerner **ét** Ctrl+Z det hele inklusive markeringen, og Ctrl+Y laver det om igen; panelet viser linjen om Ctrl+Z
- [ ] Samme i kodetilstand: et script med flere `context.sync()` er ét Ctrl+Z
- [ ] Uden ExcelApi 1.20 (Excel 2021/2024 LTSC, Excel 2019, ældre Microsoft 365): planen udføres uden fejl, linjen om Ctrl+Z vises ikke, og **Fortryd agent-session** gendanner alt som før
- [ ] Excel på web: notér, om ExcelApi 1.20 meldes understøttet, og om ét Ctrl+Z virker
- [ ] Ctrl+Z efter udførelsen og derefter **Fortryd agent-session**: notér, hvad der sker (BRUG.md fraråder kombinationen; især slettede rækker kan blive sat ind to gange)
- [ ] Markering: de ændrede celler og nye tabelrækker bliver lysegule; **Fjern markering** giver celler med fyld deres farve tilbage, og celler uden fyld står igen uden fyld med synlige gitterlinjer (ikke hvide); fed, talformat og kanter er urørte
- [ ] Markering i en Excel-tabel med stribede rækker: efter **Fjern markering** vises tabeltypografiens striber igen
- [ ] Markering på celler med mønsterfyld eller temafarve: farven er den samme efter **Fjern markering** (gradueringer gendannes ikke; notér det, hvis det ses)
- [ ] **Fortryd agent-session** uden at fjerne markeringen først: ingen gule celler tilbage, og formaterne er som før planen
- [ ] En ny plan, mens den gamle markering står: den gamle markering forsvinder, og Fortryd for den nye plan efterlader ingen gule celler
- [ ] En plan med både skrivning og slettede rækker (eller sortering) på samme ark: intet markeres, og knappen **Fjern markering** vises ikke

**Excel – formelsporing (#76)** (Excels svar på `getDirectPrecedents` m.fl. er kun kendt fra typerne og dokumentationen; simuleringen udleder afhængighederne af formlernes henvisninger. Lav et testark med en `#DIV/0!`-kæde over to ark, fx `Resultat!B3 =B2+1`, `Resultat!B2 =Beregning!B5*2`, `Beregning!B5 =B3/B4` og `Beregning!B4 =Data!B2-Data!B3` med ens tal i Data):
- [ ] **Find fejlkilden** på `Resultat!B3`: svaret peger på `Beregning!B5`, og et klik på adressen skifter ark og markerer cellen; intet er ændret i projektmappen, og Ctrl+Z har ikke fået et nyt trin
- [ ] Adressernes form: notér, hvad `areas.items[].address` indeholder for præcedenter på flere ark, for flere områder på samme ark (skilletegn, mellemrum), for arknavne med mellemrum eller apostrof, og om naboceller slås sammen til ét område (`parseAreas` i `excelTrace.ts` forventer arknavn pr. område og komma eller semikolon imellem)
- [ ] En formel uden henvisninger blandt præcedenterne (fx `=I DAG()` eller `=5*2`): sporingen fortsætter for de andre celler (Excel svarer `ItemNotFound` ved `context.sync()`, og add-in'et spørger igen i mindre dele); bekræft, at samme `Excel.run` kan bruges videre efter den fejlede rundtur
- [ ] En celle, der dividerer med en tom celle: den tomme celle står i listen som "(empty)", og formlen udpeges som kilden
- [ ] Hele kolonner (`=SUM(Data!B:B)`) og store områder (`=SUM(B2:B5000)`): kun den brugte del læses, svaret kommer hurtigt, og fejlceller langt nede i området findes
- [ ] Navngivne områder, tabelhenvisninger (`Tabel1[Beløb]`), `INDIRECT`/`OFFSET` og henvisninger til en anden projektmappe: notér, hvad Excel oplyser som præcedenter, og om noget står under "Excel could not trace or read" i forespørgslen
- [ ] En lang kæde (over 25 led, fx en løbende saldo ned gennem en kolonne): svaret siger, at sporingen er beskåret, og nævner den sidste celle; i Microsoft 365 står Excels samlede graf (`getPrecedents`, ExcelApi 1.14) i forespørgslen; notér ventetiden i Excel på web
- [ ] En cirkulær henvisning: sporingen stopper, og task pane hænger ikke
- [ ] **Forklar formel** på en formel med henvisninger til et andet ark: forklaringen nævner, hvor tallene kommer fra, med klikbare adresser
- [ ] Agenten: *"B3 viser #DIV/0! – find fejlen og ret den"* viser "Læser B3…", og planen retter formlen i kildecellen; *"hvad bruger Data!B2?"* giver de afhængige celler (kræver ExcelApi 1.13)
- [ ] Excel 2019 eller ældre (uden ExcelApi 1.12): **Find fejlkilden** viser forklaringen om, at Excel ikke kan spore, **Forklar formel** virker som før, og agenten læser områder i stedet for at spore
- [ ] Excel 2021 (ExcelApi 1.12–1.14 uden 1.15): sporing af både præcedenter og afhængige virker; notér, hvilke requirement sets der meldes understøttet på Windows, Mac og web

**Agent i PowerPoint – formatering** (præsentation med firmaskabelon, fed ord i en punktliste, farvet tekst):
- [ ] Erstat et tal i et punkt med fed tekst foran: kun tallet skifter, fed og farve bevares
- [ ] Ny tekst i en pladsholder: skabelonens skrifttype, størrelse, farve og punkttegn bevares; fed på markerede ord
- [ ] "Tilføj et slide": oprettes med et layout fra skabelonen og står det rigtige sted; teksten står i pladsholderne
- [ ] Slet og flyt slides (Microsoft 365), og fortryd: slides er tilbage på deres plads med al formatering (billeder, noter, animationer)
- [ ] PowerPoint 2016/2019: sletning og flytning afvises med en forklaring; fortryd gendanner teksten

**Design: master og layout (PowerPoint)**, prøv alle tre stile med en præsentation med tekst og tal:
- [ ] Forhåndsvisningen svarer til resultatet (farver, placering, skrifttyper)
- [ ] Under Vis → Diasmaster findes én ny master med layoutene "‹navn› – Titel/Punkter/…" og deres pladsholdere
- [ ] Nyt slide → layoutene vises og kan bruges; agentens "addSlide" bruger dem
- [ ] Tekst i pladsholderne har designets skrifttype, størrelse og farve; punkter har punkttegn
- [ ] "Erstat de nuværende slides" + Fortryd: de gamle slides er tilbage i rækkefølge med billeder, noter og animationer
- [ ] Aptos uden Microsoft 365 (fx PowerPoint 2019): erstattes af en læselig skrifttype
