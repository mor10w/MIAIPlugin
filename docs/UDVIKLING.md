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
Bridgen kører på `https://localhost:44321` og serverer også task pane. Er porten optaget (fx af en anden brugers bridge), vælges og huskes den næste ledige. `MIAI_PORT` låser porten.

API'et kræver et adgangstoken pr. bruger (#11). Bridgen skriver dit personlige manifest med port og token i konfigurationsmappen og viser stien ved start. Sideload det:
```bash
npm run sideload:mac -- ~/Library/Application\ Support/MIAI/manifest.xml
```
På Windows: `npm run sideload:windows -- %APPDATA%\MIAI\manifest.xml`. Linux: `~/.config/miai/manifest.xml`.

Watch-tilstand for både add-in og bridge: `npm run dev:bridge`.

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

**Udviklerudgave (ADR-0016, tillæg):** hvert push til `main`, der rører appen, bygger en `.dmg` til Apple Silicon i `dev-build.yml` med version `<version>-dev.<kørselsnummer>` og lægger den som forhåndsudgivelsen `dev`. Pages serverer den under `<pages-url>/dev/`, og appens menu viser *Hent udviklerudgave …*, når der er en nyere. Workflowet kan også startes manuelt (Actions → Udviklerudgave); fra andre grene end `main` bygges der kun, og `.dmg`'en ligger som artefakt på kørslen. Tjekliste på en rigtig Mac:
- Menupunktet dukker op i en ældre udgave og åbner hentningen i browseren.
- Appen kan godkendes i Systemindstillinger → Anonymitet & sikkerhed → *Åbn alligevel* og starter bagefter.
- Den nye udgave viser *Udviklerudgave X er installeret*, og nøgler i nøgleringen og Office-installationen virker stadig.

Lokalt: `npm run dist:mac -w @miai/app` eller `npm run dist:win -w @miai/app`, usigneret med mindre et certifikat findes. Release-workflowet kan også startes manuelt (Actions → Release); så bygges der kun, uden release.

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

- **Simuleret Office.js** (`e2e/fixtures/office.ts`): Word, Excel og PowerPoint med de API'er, adapterne bruger. Kalder en adapter et API, simuleringen ikke kender, fejler det med "Simuleret Office.js mangler: …" – udvid så simuleringen. Alt, adapterne gør ved dokumentet, registreres og kan efterprøves (`tp.calls()`, `tp.find("picture")`).
- **Mock-model** (`e2e/fixtures/servers.ts`): OpenAI-kompatibel server med streaming, ikke-streaming og værktøjskald, med eller uden CORS. Alle forespørgsler gemmes (`tp.model.lastRequest`).
- **Tilstande:** direkte (add-in hostet statisk over HTTPS) og bridge (den rigtige, byggede bridge med midlertidig konfiguration). Alle porte er dynamiske, så en lokal Ollama ikke forstyrrer.
- Testene fejler også, hvis task pane kaster en uventet JavaScript-fejl.

Lokalt bruges den Chromium, Playwright har installeret (`npx -w @miai/addin playwright install chromium`).

Browsertestene erstatter ikke den manuelle test i rigtig Office (testlisten nedenfor). De fanger regressioner i UI og adaptere, før den manuelle runde.

### Demovideo
`packages/addin/e2e/demo/` optager en kort demovideo af task pane ved siden af et simuleret dokument: korrektur som sporede ændringer, agenten med plan og godkendelse, og chat i Word, og derefter agenten i Excel (tomme rækker, tabelformat, datalinjer og diagram). Den bruger den simulerede Word og Excel fra browsertestene og indspillede modelsvar, så optagelsen er ens fra gang til gang og kan laves om, når brugerfladen ændrer sig. Den viser altså brugerfladen, ikke en models kvalitet eller hastighed, og scenen siger det selv i øverste hjørne.

```bash
npm run build
```
```bash
npx playwright test -c packages/addin/e2e/demo/playwright.config.ts
```
Optagelsen (1920×1080, ca. halvandet minut) ender i `packages/addin/demo-results/`. Til MP4:
```bash
ffmpeg -i packages/addin/demo-results/demo.video.ts-demovideo/video.webm -c:v libx264 -crf 20 -pix_fmt yuv420p -movflags +faststart -an miai-demo.mp4
```
Dokumentet, regnearket, modelsvarene og rækkefølgen står øverst i `demo.video.ts`; scenen (dokument, regneark, billedtekster, markør) er `stage.html`. `MIAI_DEMO_PROBE=1` kører det hele igennem på få sekunder og gemmer et stillbillede pr. scene.

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
- [ ] Opdatering: installér en ældre version, udgiv en nyere → menuen viser "Genstart og opdater til …", og den nye version kører bagefter
- [ ] Signeret build (når Azure Trusted Signing er sat op): ingen SmartScreen-advarsel; udgiver vises korrekt

**MIAI App (Mac)**
- [ ] "Start automatisk ved login" til: MIAI står under Systemindstillinger → Generelt → Login-emner og starter efter log ud/ind (én instans, ingen dubletter)
- [ ] Slå det fra igen: appen bliver kørende og starter ikke ved næste login
- [ ] Startet fra DMG'en (ikke i Programmer): appen tilbyder at flytte sig til Programmer
- [ ] Opgradering fra en udgave med LaunchAgent (`~/Library/LaunchAgents/dk.miai.app.plist`): filen fjernes, login-elementet er slået til, og appen starter kun én gang ved login
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
- [ ] Kodetilstand i Word: "Skriv 'Fortroligt' i sidehovedet" kører; et script, der fejler midtvejs, efterlader brødteksten som før; Fortryd gendanner brødteksten med formatering, billeder og tabeller
- [ ] Kodetilstand i PowerPoint (Microsoft 365): et script, der tilføjer en figur eller sletter slides, kan fortrydes, og alle slides er tilbage med billeder, noter og animationer; i PowerPoint 2019 afvises scriptet med en forklaring
- [ ] Politik `agentScripts: false` (bridge-politik eller `miai-config.json`): kontakten er slået fra og kan ikke ændres

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
