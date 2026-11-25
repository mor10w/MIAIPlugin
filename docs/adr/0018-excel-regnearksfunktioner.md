# ADR-0018: Regnearksfunktionerne =MIAI.ASK() og =MIAI.MODELS() i Excel

- **Status:** Besluttet; slået til i alle manifester til Word, Excel og PowerPoint. Afprøvet i rigtig Excel på Mac gennem bridgen, ikke på Windows, på web eller med det hostede manifest (#80)
- **Dato:** 2026-11-25
- **Relateret:** ADR-0007 (direkte forbindelse og bridge), ADR-0008 (sikkerhed og politik), ADR-0012 (agent-tilstand), ADR-0015 (central udrulning)

## Kontekst

Formler kan ikke kategorisere fritekst, trække et beløb ud af en beskrivelse eller oversætte en kolonne. En funktion i cellen kan, og den kan trækkes ned over tusind rækker. Microsofts `=COPILOT()` kræver Copilot-licens og har en kvote. MIAI kan tilbyde det samme med brugerens egen model.

Excel har et API til egne funktioner (custom functions). Det rejser fem valg.

## Beslutning

### 1. Shared runtime, kun for Excel

Egne funktioner kører som standard i en ren JavaScript-runtime uden `localStorage` og DOM. Dér kan de ikke nå MIAI's indstillinger, politik eller bridge-token. Med shared runtime kører de i samme side som task pane og går gennem den samme backend som opgaverne. `allowedHosts`, CSP og audit-loggen gælder derfor uændret, og der er ingen ny vej ud af maskinen.

- `<Runtimes>` og `CustomFunctions`-udvidelsen står kun under `<Host xsi:type="Workbook">`.
- Med shared runtime har Excels `<Action>` intet `<TaskpaneId>`.
- **Intet `SharedRuntime`-krav øverst i manifestet.** Manifestet deles af Word, Excel og PowerPoint, og kravet ville skjule MIAI i Word og PowerPoint på klienter uden shared runtime (iPad, volumenlicenseret Office).
- I en Excel uden shared runtime indlæser Excel `functions-fallback.js`, som svarer `#N/A` med en forklaring i stedet for at hænge.

### 2. En almindelig asynkron funktion, ikke streaming

Svaret streames ikke ind i cellen. Hver delvis tekst ville være en ny celleværdi, så afhængige formler regner om for hvert stykke, og en kæde af `MIAI.ASK`-celler ville starte og afbryde modelkald mange gange. Det samme gælder en pladsholder som "Venter…". Excel viser selv `#BUSY!`, og afhængige celler venter.

### 3. To profiler i ét argument

`=MIAI.ASK(prompt; [input]; [optimized]; [model])`.

| | `optimized = FALSK` (standard) | `optimized = SAND` |
|---|---|---|
| Websøgning og søgning i egne mapper | Som under Indstillinger og politikken | Fra |
| Tænkning | Som under Indstillinger | Fra |
| Systemprompt | Normal | Kortest mulig |
| Svarlængde | Indstillingen | 256 tokens |

Standarden er den fulde profil, fordi websøgning i en celle er nyttig (opslag) og ellers ville være umulig. Prisen er, at en kolonne uden `optimized` kan give mange søgninger; køen og cachen begrænser det, og vejledningen anbefaler `SAND` til mange rækker. Begge profiler bruger temperatur 0 og en fast `seed`.

`model` vælger en model på den aktive endpoint. Den kan ikke pege på en anden server. Et ukendt navn giver `#VALUE!` i stedet for et stille fald tilbage til standardmodellen.

### 4. Kø, cache og talregel i add-in'et

- **Kø (`cell/queue.ts`):** et loft over samtidige kald (indstillingen `cellConcurrency`, standard 2). Celler med samme forespørgsel deler ét kald. Køen kan stoppes fra task pane.
- **Cache (`cell/cache.ts`):** nøglen er en SHA-256-hash af endpoint, model, profil, prompt og input. Cachen gemmes i `localStorage` med et loft og kan ryddes. Den gemmer kun hash og svar. Fejl caches ikke. Cachen er nødvendig, fordi Excel kalder funktionen igen ved genberegning og muligvis ved åbning.
- **Talregel (`cell/value.ts`):** svaret bliver kun til et tal, når hele svaret er et entydigt tal uden foranstillet nul. "0045" og "1.234" forbliver tekst.
- **Audit-log:** i direkte tilstand samles en kørsels cellekald i én post (`mergeCellAudit`), fordi loggen dér kun rummer 200 poster. Bridgen logger hvert kald.

### 5. Modelvalg i en celle

Excel kan kun vise faste forslag til et argument, og kun for tekst og tal. Brugerens modeller kan derfor ikke foreslås, mens formlen skrives. `=MIAI.MODELS()` fylder i stedet en kolonne med modellerne, og knappen "Indsæt modelvælger" laver en dropdown (datavalidering), som peger på listen på et skjult hjælpeark.

## Udrulning i to trin

Et manifest med en skemafejl indlæses slet ikke, heller ikke i Word og PowerPoint, og Microsofts validator kan ikke nås fra byggemiljøet. Derfor:

1. Koden er med i alle udgaver, men `scripts/make-manifest.mjs` lægger kun shared runtime og funktionerne i `manifest.dev.xml`. Et manifest med funktionerne åbner task pane med `?functions=1`, og kun da registreres funktionerne, og statuslinjen og sektionen under Indstillinger vises. Siden kan ikke se det på `CustomFunctions`: Excels office.js definerer objektet i enhver task pane, også uden funktionerne i manifestet. Den første udgave (0.1.23) antog det modsatte og viste derfor sektionen i de udgivne installationer, hvor funktionen ikke fandtes i arket.
2. Når udviklingsmanifestet er afprøvet i rigtig Excel på Windows og Mac efter tjeklisten i [UDVIKLING.md](../UDVIKLING.md), slås funktionerne til i bridge-manifestet og det hostede manifest. Bridgens personlige manifest (`personalizeManifest`) sætter allerede sit token på en adresse med `?functions=1`.

**Trin 2 er taget den 25. november 2026**, før betingelsen var helt opfyldt: funktionerne var afprøvet i Excel på Mac gennem bridgen med et manifest, der er identisk med det nu udgivne bridge-manifest, men ikke på Windows, ikke i Excel på web og ikke med det hostede manifest. Tjeklisten i UDVIKLING.md står derfor stadig åben for dem. Går et manifest i stykker dér, slås funktionerne fra igen for den variant med `functions: false` i `scripts/make-manifest.mjs`.

Under afprøvningen fandtes testmanifestet `manifest.functions.xml`: bridge-manifestet med funktionerne slået til. Det er fjernet igen, fordi bridge-manifestet nu selv har dem.

## Konsekvenser

- Udgivne installationer får funktionerne med et nyt manifest: MIAI App lægger det selv ind ved næste opdatering, mens det hostede manifest skal hentes og lægges ind igen (eller uploades igen ved central udrulning). Indtil da ændres intet for dem.
- Med shared runtime lever siden, så længe projektmappen er åben. Genindlæsningen ved ændret politik (#28) afbryder kørende funktioner.
- En projektmappe med `MIAI.ASK` viser den sidst beregnede værdi hos en kollega uden MIAI, og `#NAME?` ved genberegning.

## Erfaringer fra rigtig Excel (Mac, november 2026)

- Manifestet indlæses, task pane åbner uden `<TaskpaneId>`, og `=MIAI.ASK(` findes i arket med testmanifestet gennem bridgen.
- Cellen stod på `#BUSY!` uden noget kald til bridgen. Funktionen ventede på panelets opstart, som kun afsluttes, når `Office.onReady` melder Excel som vært. Funktionerne opretter derfor selv deres backend, så snart siden er indlæst, og knyttes til, så snart Office.js har defineret `CustomFunctions`. Om Excel på Mac kører funktionerne i samme side som panelet eller i sin egen skjulte udgave af den, er ikke afklaret; koden virker nu i begge tilfælde.
- Et kald venter højst 20 sekunder på MIAI selv (backend, indstillinger, modelliste). Derefter viser cellen `#N/A` med en forklaring.
- Kører funktionerne i en anden udgave af siden end panelet, kan panelets statuslinje og Stop-knap ikke se køen. Det skal afklares.
- Med testmanifestet gennem bridgen giver `=MIAI.ASK("…"; B3; SAND)` svar i cellen (Excel 16.115 på Mac, oMLX). Det første kald tog omkring 40 sekunder, mens modellen blev indlæst; derefter kom svarene hurtigt.
- `#BUSY!` uden kald til bridgen havde endnu en årsag, som ikke ligger i koden: Excel gemmer, hvilket add-in navnerummet `MIAI` hører til. Maskinen havde først brugt udviklingsmanifestet og fik derefter testmanifestet, som har et andet id, men samme navnerum. Excel hentede testmanifestets `functions.json`, men beholdt den gamle registrering, og formlen stod på `#BUSY!` eller var ukendt i arket. Efter at registreringen var slettet med Excel lukket, oprettede Excel en ny ved næste start, og funktionerne virkede. Fremgangsmåden står i [UDVIKLING.md](../UDVIKLING.md#skift-mellem-manifester-med-funktionerne).
- Det har en følge for trin 2 under Udrulning i to trin: en bruger, der har haft et andet MIAI-manifest med funktionerne, kan ramme det samme, når funktionerne slås til i bridge-manifestet og det hostede manifest. Efter trin 2 gælder det enhver, der skifter mellem MIAI App, det hostede add-in og udviklingsmanifestet på samme maskine.

## Skal bekræftes i rigtig Excel

- At manifestet indlæses i Word, Excel og PowerPoint.
- Om funktionerne kaldes igen, når en gemt projektmappe åbnes.
- At datavalidering sat fra Office.js accepterer en henvisning til et dynamisk område (`=MIAI_modeller!$A$1#`).
- At kald til localhost virker fra shared runtime på Windows (loopback, #9 og #68).
- At task pane åbner som før i Excel uden `<TaskpaneId>`.
