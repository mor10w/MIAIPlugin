# ADR-0017: Flere faste værktøjer og en kodetilstand for agenten (Excel)

- **Status:** Besluttet (prototype i Excel; afprøves i rigtig Excel, før den udvides til Word og PowerPoint)
- **Dato:** 2026-11-04
- **Relateret:** ADR-0012 (agent-tilstand), ADR-0008 (sikkerhed og politik)

## Kontekst

Agenten kan kun det, der er bygget et værktøj til. Brugerne beder om meget mere, end de faste værktøjer dækker, fx sortering, diagrammer, pivottabeller og datavalidering. Hver gang agenten ikke har et værktøj, svarer den med en forklaring i stedet for at gøre det.

Tre måder at give agenten adgang til "alle funktioner" blev overvejet:

| Vej | Fordel | Ulempe |
|---|---|---|
| **Styre brugerfladen** (skærmbilleder, mus og tastatur via OS-automatisering eller Playwright på Office på web) | Alt, hvad en bruger kan klikke på | Kræver en vision-model; langsom og skrøbelig over for Office-version, sprog og layout; ingen plan at godkende og intet fortryd; den største sikkerhedsrisiko. Kan ikke køre i et add-in. |
| **COM/AppleScript fra MIAI App** | Næsten hele VBA-objektmodellen | Kun Windows i fuldt omfang, ikke Office på web; makro-magt, som er svær at afgrænse |
| **Modellen skriver Office.js-kode**, som add-in'et kører | Hele Office.js; virker på Windows, Mac og web; passer til add-in'et | Koden er modeloutput (prompt injection) og skal afgrænses; fortryd er kun omtrentlig; kræver en stærk model |

Et Office-add-in kan ikke klikke i Office' bånd eller se skærmen. Det har kun Office.js. Det gælder også andre AI-add-ins til Office.

## Beslutning

To spor:

### 1. Flere faste værktøjer (altid til rådighed)

Excel-agenten får `sortRange`, `conditionalFormat`, `addChart` og `freezeRows`, ud over `deleteRows`, `deleteEmptyRows`, `formatTable`, `setNumberFormat` og `autofitColumns` (ADR-0012).

- **Ingen fri formatering:** betinget formatering bruger Excels standardforvalg (rød-gul-grøn farveskala, datalinjer og "lys rød udfyldning med mørkerød tekst"), og diagrammer bruger temaets farver.
- **Præcis fortryd:**
  - sortering: cellerne gendannes fra et øjebliksbillede;
  - diagrammer og betinget formatering: de oprettede slettes (efter navn og id);
  - frysning: den tidligere frysning sættes tilbage.
- **Krav:** betinget formatering kræver ExcelApi 1.6, og frysning kræver ExcelApi 1.7. Mangler de, meldes det pr. handling.

De faste værktøjer foretrækkes altid. De fungerer også med små modeller, og de kan fortrydes præcist.

### 2. Kodetilstand: `runScript {description, code}` (slået fra som standard)

Til det, de faste værktøjer ikke kan, fx pivottabeller, datavalidering, navngivne områder og sortering på flere kolonner.

- **Slås til** under Indstillinger (`settings.agentScripts`). Politikken kan forbyde det (`policy.agentScripts: false`), også centralt via `miai-config.json` (ADR-0015).
- **Kun når tilstanden er slået til**, får modellen vejledningen til `runScript`. Motoren afgør det (`params.agentScripts`), så bridge og direkte tilstand er ens. `checkPlan` fravælger et script, når tilstanden er slået fra (`scripts-off`).
- **Scriptet skal stå alene** i planen (`with-script`), fordi fortryd gendanner arket fra lige før scriptet.
- **Brugeren ser beskrivelsen og kan se koden**, før noget køres. Scriptet er markeret *Fjerner indhold*.

**Sandkassen** (`shared/src/script.ts`, `addin/src/script.ts`):
- **Statisk kontrol med acorn (kapabilitetsbaseret):**
  - Hvert navn, scriptet bruger, skal være erklæret i scriptet med korrekt blok-scope, være en parameter (`context`, `Excel`) eller stå på listen over sikre globale objekter (`Math`, `JSON`, `Array` …).
  - Alle andre globale objekter afvises, også dem, vi ikke kender: `window`, `fetch`, `localStorage` (med indstillinger og API-nøgler), `document`, `eval`, `Function`, `setTimeout` osv.
  - Spærret er også: `this`, klasser, `import()`, `new.target`/`import.meta`, `with` og navne, der starter med `__`.
- **Egenskaber, der kan føre ud af sandkassen, er spærret:**
  - fx `constructor`, `prototype`, `__proto__`, `getPrototypeOf`, `defineProperty` og `prepareStackTrace`;
  - interne Office.js-felter (`_…`, `$…`, `m_…`);
  - også i destrukturering og objekt-literaler.
- **Beregnede navne (`obj[x]`) tjekkes ved kørsel:** koden omskrives til `obj[__key(x)]`, og `__key` gør navnet til en primitiv værdi, før det tjekkes.
- **Løkkevagt og tidsgrænse:** hver løkke får `__tick()` (højst 1.000.000 gennemløb i alt), og scriptet har en tidsgrænse på 60 sekunder.
- **Kørsel:** strict mode (`new Function`) med kun `context`, `Excel`, `__key` og `__tick`.

**Øjebliksbillede, tilbagerulning og fortryd:**
- **Før kørslen gemmes:** det aktive arks brugte område (formler, talformater og celleegenskaber) samt navnene på projektmappens ark, tabeller og pivottabeller og arkets diagrammer.
- **Ark over 50.000 celler afvises**, for så kan ændringerne ikke fortrydes.
- **Fejler scriptet, rulles ændringerne straks tilbage.** Fejlen med Office' fejlplacering (`debugInfo.errorLocation`) og koden kan sendes tilbage til modellen med **Bed modellen rette scriptet**.
- **Fortryd fjerner det, scriptet har tilføjet**, og gendanner arkets celler: nye ark, pivottabeller og diagrammer slettes, og nye tabeller bliver almindelige områder igen.
- **Ikke dækket:** ændringer på andre ark, slettede ark, betinget formatering og datavalidering. Det står i brugerfladen.

## Konsekvenser

- Agenten kan nu det meste, Excel kan, uden at der skal bygges et værktøj til hver funktion. De almindelige ønsker klares stadig med faste værktøjer med præcis fortryd.
- **Sikkerhed:** sandkassen er et forsvar i dybden, ikke en garanti. Tre ting hjælper:
  - brugeren godkender hvert script og kan se koden;
  - politikken kan slå tilstanden fra;
  - tilstanden er slået fra som standard.

  **Næste skridt:** en Content-Security-Policy på task pane (`connect-src` begrænset til de tilladte hosts), så netværk også er spærret uden for sandkassen.
- **Modelkrav:** små modeller (4B) skriver ofte fejlbehæftet Office.js. Tilbagerulningen og muligheden for at lade modellen rette scriptet gør fejlene ufarlige, men tilstanden giver mest med 14B+ eller en kodemodel. Eval-casen `agent-excel-script-pivot-da` måler det (#2).
- **Mangler:** Word og PowerPoint. Samme mønster kan bruges, med dokumentets OOXML (Word) eller `exportAsBase64` (PowerPoint) som øjebliksbillede.

**Test:**
- Unit-tests af sandkassen: globale objekter, prototyper, blok-scope, omskrivning og løkker.
- Unit-tests af `checkPlan`.
- E2e i simuleret Excel:
  - faste værktøjer med fortryd;
  - kodetilstand slået fra;
  - et script, der køres og fortrydes;
  - netværk afvist;
  - et fejlende script rulles tilbage og rettes af modellen.
- Tjekliste til rigtig Excel i UDVIKLING.md.

## Tillæg: pivottabel og datavalidering som faste værktøjer (#25)

Pivottabeller og datavalidering var de to mest almindelige grunde til at slå kodetilstanden til. De er nu faste værktøjer, så de virker med små modeller og med præcis fortryd:

- **`addPivotTable {source, rows, columns?, values, destination}`** (ExcelApi 1.8): kilden er en tabel eller et område med overskrifter, felterne navngives med overskriftsteksten, og hver værdi kan opsummeres med sum, antal, gennemsnit, største eller mindste. Pivottabellen hører til destinationens ark og kan lægges på et ark, planen selv opretter.
- **`setValidation {address, list | min/max, whole?}`** (ExcelApi 1.8): rulleliste med tilladte værdier eller et tal mellem to grænser.
- **Validering før udførelse:** felter, der ikke er overskrifter i kilden, fravælges (`bad-field`), når overskrifterne kendes (tabeller og områder fra arkets første brugte række). En regel skal være enten liste eller talgrænse, og listeværdier må ikke indeholde komma.
- **Udførelse én ad gangen:** hver pivottabel og regel synkroniseres for sig. Afviser Excel en af dem, fejler kun den handling, og en halvt oprettet pivottabel slettes igen.
- **Fortryd:** pivottabellerne slettes, og området får sin tidligere valideringsregel tilbage (eller ingen).

Kodetilstanden er fortsat til det, de faste værktøjer ikke dækker, fx navngivne områder, sortering på flere kolonner og filtre i pivottabeller.

## Tillæg: Content-Security-Policy på task pane (#28)

Sandkassen beskytter mod kode fra modellen. CSP'en er laget udenom: browseren håndhæver selv, hvor task pane må sende data hen, uanset hvordan koden er kommet ind.

- **Direktiver:** `connect-src` (fetch, XHR, WebSocket, sendBeacon), `img-src 'self' data: blob:`, `media-src`, `object-src 'none'`, `form-action 'none'` og `base-uri 'self'`.
- **Direkte tilstand:** `connect-src` er add-in'ets egen adresse plus politikkens hosts (http og https på alle porte; `*.domæne` og IPv6 understøttes). CSP'en sættes som `<meta>` ved start, før noget vises. Navne, der ikke kan udtrykkes sikkert i en CSP, udelades.
- **Bridge:** task pane taler kun med bridgen, så `connect-src 'self'`. Bridgen sender samme CSP som header, og add-in'et sætter den også selv.
- **Ændret politik:** en CSP kan kun strammes, mens siden er åben. Fjernes hosts, strammes den med det samme. Tilføjes hosts, genindlæses task pane, åbner Indstillinger igen og siger, at forbindelsen kan testes.
- **Bevidst ikke `script-src`/`default-src`:** Office.js hentes fra Microsofts CDN og indlæser selv mere kode, og kodetilstanden bruger `new Function`. En stram script-politik kan ikke afprøves i rigtig Office herfra; målet er at spærre dataudslip.
- **Dækker ikke:** navigation væk fra siden (`window.location`) og Office.js' egne API'er. Office.js' telemetri til Microsoft bliver blokeret af `connect-src`; det er uden betydning for funktionen, men skal bekræftes i rigtig Office (tjeklisten i UDVIKLING.md).
