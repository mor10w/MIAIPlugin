# ADR-0018: Regnearksfunktionerne =MIAI.ASK() og =MIAI.MODELS() i Excel

- **Status:** Besluttet; kun slået til i udviklingsmanifestet og kun afprøvet i simuleret Office (#80)
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

Til afprøvning med bridgen findes testmanifestet `manifest.functions.xml`: bridge-manifestet med funktionerne slået til. Det bruges med bridgen fra kildekoden (`MIAI_MANIFEST`, se [UDVIKLING.md](../UDVIKLING.md)) og er nødvendigt, når modelserveren ikke tillader CORS, fordi udviklingsmanifestet kører direkte mod modellen. MIAI App bruger det ikke.

## Konsekvenser

- Ingen ændring for udgivne installationer, før trin 2.
- Med shared runtime lever siden, så længe projektmappen er åben. Genindlæsningen ved ændret politik (#28) afbryder kørende funktioner.
- En projektmappe med `MIAI.ASK` viser den sidst beregnede værdi hos en kollega uden MIAI, og `#NAME?` ved genberegning.

## Skal bekræftes i rigtig Excel

- At manifestet indlæses i Word, Excel og PowerPoint.
- Om funktionerne kaldes igen, når en gemt projektmappe åbnes.
- At datavalidering sat fra Office.js accepterer en henvisning til et dynamisk område (`=MIAI_modeller!$A$1#`).
- At kald til localhost virker fra shared runtime på Windows (loopback, #9 og #68).
- At task pane åbner som før i Excel uden `<TaskpaneId>`.
