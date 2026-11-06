# ADR-0012: Agent-tilstand – plan som struktureret output, udført i add-in'et

- **Status:** Besluttet
- **Dato:** 2026-11-03
- **Issue:** #6

## Kontekst

Brugeren skal kunne beskrive en ændring ("ret alle 'kunde' til 'klient' og tilføj en overskrift"), hvorefter MIAI selv ændrer dokumentet. Kravene fra PLAN.md §6 og #6:

- Modellen må kun bruge whitelistede Office-handlinger.
- Hver handling skal bekræftes: **plan → bekræft → udfør**.
- Dokumentindhold kan indeholde prompt injection.
- Hele agent-sessionen skal kunne fortrydes.

Office.js findes kun i add-in'ets webview, ikke i bridgen. Klassisk tool calling (modellen kalder et værktøj, får resultatet og fortsætter) ville derfor kræve en rundtur fra bridgen til add-in'et og tilbage for hvert kald, mens motoren venter. Det er også i strid med "bekræft før udfør": modellen forventer et resultat, før brugeren har set planen.

## Beslutning

1. **Planen er struktureret output, ikke tool calls.** Agenten er en almindelig opgave (`agent`) med JSON-skemaet `AgentPlan`: en kort opsummering og højst 20 handlinger fra en lukket union (`AgentAction`).
   - Word: `replace`, `insert`, `delete`, `move`, `comment`.
   - Excel: `setValues`, `setFormula`, `clear`.

   Handlingerne svarer til tool calls i én runde, men de går gennem motorens eksisterende vej for struktureret output (`json_schema` → `json_object` → prompt + validering + ét genforsøg). Det virker derfor også med værktøjer uden tool calling, og ukendte handlinger afvises af skemaet.
2. **Motoren udfører intet.** Planen vises i add-in'et. `checkPlan` markerer, inden brugeren ser planen, handlinger, der ikke kan udføres: forkert app, tekst der ikke findes i dokumentet, ugyldig adresse eller for mange celler. De markerede handlinger er fravalgt og vises med årsagen. Destruktive handlinger (slet, ryd, erstat med tomt) fremhæves.
3. **Brugeren vælger og bekræfter.** Kun de afkrydsede handlinger udføres, og først når brugeren trykker *Udfør*. I Word kan de udføres som sporede ændringer (WordApi 1.4, slået til som standard), så de også kan gennemgås i Word.
4. **Fortryd for hele sessionen.**
   - Word: et OOXML-øjebliksbillede af dokumentets indhold tages lige før udførelsen. *Fortryd agent-session* gendanner det med `insertOoxml("Replace")`. UI'et siger tydeligt, at det også fjerner ændringer, brugeren selv har lavet siden.
   - Excel: formlerne i alle berørte områder gemmes, og ved fortryd skrives de tilbage i omvendt rækkefølge, samme mønster som `applyFormula`.
5. **Sikkerhed.**
   - Der er ingen netværks-, fil- eller formateringsværktøjer i agent-tilstand og ingen web-søgning. Agenten kører uden `ToolExecutor`.
   - Dokumentet står i `<document>`-tags, og systemprompten siger, at instruktioner derinde er data.
   - Modellen ser aldrig resultatet af sine handlinger, så en kapret plan kan ikke eskalere i flere trin. Det værste, den kan, er at foreslå tekstændringer, som brugeren ser, før de udføres.
   - Excel må kun skrive på markeringens ark, højst 2.000 celler pr. handling.

## Konsekvenser

- Samme flow i direkte tilstand og bridge-tilstand. Der er ingen ny protokol mellem bridge og add-in.
- Agenten kan ikke "kigge" i dokumentet undervejs (fx søge, læse resultatet og planlægge videre). Den arbejder på markeringen eller hele dokumentet, som add-in'et sender med (højst ca. 24.000 tegn). Flertrinsplaner med læseværktøjer kan tilføjes senere som en udvidelse, når der er behov.
- `find`-tekst skal være ordret. Små modeller rammer ikke altid præcist. Den slags handlinger fanges af `checkPlan` og vises som "Teksten findes ikke i dokumentet" i stedet for at fejle stille. Eval-sættet (`cases/agent.json`) måler, hvor ofte planerne er gyldige.
- ~~Flyt sker som tekst.~~ Erstattet af tillægget nedenfor: hele afsnit flyttes med deres formatering.

## Tillæg (2026-11-04): Formateringen bevares

Kravet er, at agenten kan redigere uden at ødelægge formateringen, og at den selv kan skabe formatering, hvor den laver nyt indhold. Den første udgave arbejdede på ren tekst og fandt alt med tekstsøgning. Det ødelagde formatering i flere tilfælde: indsat tekst arvede naboens overskrift eller liste, erstat gjorde blandet formatering ensartet, sletning efterlod tomme afsnit, og fortryd erstattede hele dokumentet.

1. **Modellen ser strukturen.** For agenten henter Word-adapteren afsnittene med typografi, liste og tabel (WordApi 1.3), og prompten viser dem som `¶3 [Heading2] Priser`, `(list)` og `(table)`.
2. **Strukturhandlinger peger på afsnit.** `insert`, `delete {paragraph}`, `move` og den nye `setStyle` bruger ¶-numre i stedet for tekstsøgning.
   - Numrene er dem, modellen så. Afsnittene hentes som objekter, før første handling udføres, så numrene gælder, selv om tidligere handlinger har indsat eller slettet afsnit.
   - Er et afsnit ændret mellem plan og udførelse, afvises handlingen ("dokumentet er ændret").
3. **Erstat ændrer kun det, der er forskelligt** (`minimalEdit`): fælles ord før og efter bevares, og kun midten skiftes ud. Words regel om, at erstattet tekst får første tegns formatering, rammer dermed kun de ændrede ord.
4. **Ny tekst arver ikke naboens formatering.**
   - Hver linje får en eksplicit indbygget typografi.
   - Arvet direkte fed, kursiv og understregning nulstilles for brødtekst og lister.
   - Arvet listemedlemskab fjernes (`detachFromList`), når typografien ikke er en liste.
   - Begrænset Markdown giver overskrifter, punkter, nummererede lister, citater, fed, kursiv og links. Kun http(s) og mailto er tilladt. Skrifttype, farve og størrelse kommer altid fra dokumentets typografier, så der stadig ikke er fri formatering.
5. **Hele afsnit slettes og flyttes som afsnit.** Flyt bruger afsnittets OOXML inklusive afsnitsmærket. Sletning af al teksten i et afsnit sletter afsnittet. Tabelafsnit må kun få deres tekst rettet (`checkPlan` → `in-table`).
6. **Fortryd afviser sessionens sporede ændringer** (WordApi 1.6, tidsvindue for sessionen) og fjerner agentens kommentarer via deres id. Resten af dokumentet, også brugerens egne sporede ændringer, røres ikke. Øjebliksbilledet af hele dokumentet bruges kun uden sporede ændringer eller i Word før 1.6, og UI'et siger det.
7. **Chat → agent:** chatsvaret følger med som `<material>`, så agenten indsætter den tekst, brugeren har set.

**Test:** den simulerede Word i e2e har en dokumentmodel, der efterligner Words regler: arv ved indsættelse, første tegns formatering ved erstatning og sporede ændringer. Testene tjekker dokumentets typografier og formatering efter hver handling og efter fortryd. Reglerne skal stadig bekræftes i rigtig Word; tjeklisten står i UDVIKLING.md.

### Excel og PowerPoint

Samme principper: modellen ser strukturen, handlingerne ændrer kun det nødvendige, nyt indhold får dokumentets egen formatering, og fortryd gendanner også formateringen.

**Excel**
- **Kontekst:** arkets Excel-tabeller (navn, område, overskrifter) og kolonnernes talformat i første datarække.
- **`appendRows {table, values}`** bruger `table.rows.add`, så rækkerne får tabellens formatering og beregnede kolonner (`null` = lad Excel udfylde).
- **Værdier som JSON-tal**, ikke tekst, så tal ikke tolkes efter landestandard.
- **Naboformat:** celler, der var tomme, får formatet fra rækken over eller kolonnen til venstre (`copyFrom(…, formats)`, ExcelApi 1.9), før værdien skrives.
- **Udseende uden fri formatering** (tilføjet efter brugerfeedback: "gør tabellen pæn" gav kun en vejledning):
  - `formatTable {target, style}` giver en eksisterende tabel en ny typografi eller gør et område til en Excel-tabel (`tables.add`). Kun Excels indbyggede tabeltypografier (`TableStyleLight1–21`, `Medium1–28`, `Dark1–11`) er tilladt. De følger projektmappens tema, så modellen aldrig vælger farver eller skrifttyper selv.
  - `setNumberFormat {address, format}` og `autofitColumns {address}`.
- **Sletning af rækker:**
  - `deleteRows {address}` sletter rækker. Inde i en Excel-tabel sker det som tabelrækker, så tabellen krymper. Ellers slettes cellerne i områdets kolonner, og cellerne under rykker op; data til højre for området røres ikke.
  - `deleteEmptyRows {target}` (tabelnavn eller område): add-in'et finder selv de helt tomme rækker. Små modeller tæller ofte forkert, når de skal angive rækker. Et område, der rammer en tabel, behandles som tabellen, og i et område tælles kun rækker inden for arkets data.
  - Sletninger udføres **til sidst og nedefra**, så alle adresser i planen gælder arket før handlingerne. Hver række slettes for sig, så en afvisning fra Excel ikke efterlader en halvt udført sletning uden fortryd.
  - `checkPlan` afviser sletninger, der rammer en tabels overskrift eller stikker ud af tabellen, og nye tabeller, der overlapper en eksisterende (`in-table`).
- **Rækkenumre i konteksten:** de markerede værdier vises med arkets rækkenummer først, så modellen kan pege på rækker uden at tælle.
- **Fortryd** i omvendt rækkefølge:
  1. Slettede rækker sættes ind igen (`rows.add(index, …)` i tabeller, ellers `insert(down)`) med formler, talformat og celleegenskaber.
  2. Nye tabeller bliver almindelige områder igen (`convertToRange` og ryd formater), og tidligere tabeltypografier sættes tilbage.
  3. Formler, talformat og celleegenskaber (`getCellProperties`/`setCellProperties`, ExcelApi 1.9, nu også kanter) gendannes, og det samme gør kolonnebredderne.
  4. Tilføjede tabelrækker slettes.
- `formatTable` på et område kræver ExcelApi 1.9, for ellers kan formateringen ikke fortrydes.

**PowerPoint**
- **Kontekst:** alle slides med layout og figurer som `2.1 [title|body|text|other] navn: tekst`, plus præsentationens layouts.
- **Handlinger:**
  - `replace` deles med Word og ændrer kun det forskellige via `TextRange.getSubstring` (PowerPointApi 1.4).
  - `setText` skriver én linje pr. afsnit; punkttegnene kommer fra figuren, og fed/kursiv sættes via delstrenge.
  - `addSlide` opretter fra et navngivet layout (ellers forrige slides layout) og udfylder pladsholderne. Der laves aldrig egne tekstbokse.
  - `deleteSlide` og `moveSlide`.
- **Fortryd** (PowerPointApi 1.8): berørte slides eksporteres (`exportAsBase64`) før første ændring og lægges tilbage med `insertSlidesFromBase64` (behold kildeformatering) på deres plads. Nye slides slettes, og flyttede flyttes tilbage.
- **Uden 1.8:** kun teksten kan gendannes, så sletning og flytning af slides afvises, og UI'et siger, at formatering kan gå tabt ved fortryd.
- **Agenten kræver PowerPointApi 1.4** (PowerPoint 2019/Microsoft 365).

Tests: simuleret Excel med formater og tabeller og simuleret PowerPoint med slides, layouts og formateret tekst. Testene tjekker formatering efter handlinger og efter fortryd. Tjeklister til rigtig Office står i UDVIKLING.md.

**Opfølgning:** flere faste Excel-værktøjer (sortering, betinget formatering, diagram og frysning) og en kodetilstand med Office.js-scripts i en sandkasse er beskrevet i [ADR-0017](0017-agent-kodetilstand.md).

## Tillæg (2026-11-06): saveSheet – et ark gemmes som .xlsx-fil

Brugeren kan bede agenten om at flytte et ark ud i en ny fil ("sæt dette ark i sin egen fil"). Det kan de faste handlinger ikke: Office.js har ingen eksport-API, og bridgen kan skrive filer – add-in'et kan det ikke. Undtagelsen er derfor dokumenteret eksplicit i stedet for at åbne generel filadgang:

1. **Én handling, kun Excel:** `saveSheet {fileName, sheet?, reason}` gemmer ét ark som en ny .xlsx-fil. Den ændrer ikke arbejdsbogen, og den er ikke destruktiv – den følger alle de øvrige handlinger gennem **plan → bekræft → udfør**.
2. **Kun bridge, og slået til af brugeren.** Direkte tilstand har ingen eksekverer, der kan skrive filer: handlingen vises aldrig i prompten uden bridge (`agentExport` injiceres kun, når `deps.tools !== undefined`), og `checkPlan` markerer den som *no-bridge*, hvis den alligevel kommer. I Indstillinger (Privatliv) kan brugeren slå eksporten til – fra som standard – og politikken kan forbyde den (`agentExport: false`), så organisationen styrer, om agenten må skrive filer.
3. **Datavejen:** add-in'et læser arket (værdier, formler med deres resultat, talformater og kolonnebredder) som JSON og sender det til bridgens `POST /api/export`, der bygger filen med exceljs. Formler skrives som `{formula, result}`, så filen viser tallene, før den åbnes i Excel.
4. **Sikkerhed:**
   - Mappen er et fast enum (`desktop`, `documents`, `downloads`) under brugerens hjem – aldrig fra modellen. Filnavnet valideres mod en streng regex (ingen stipelementer, `.xlsx`-påliggende, maks. 255 tegn).
   - Ingen overskrivning: en eksisterende fil afvises (409), så agenten aldrig skriver over en brugerfil.
   - Højst 50.000 celler pr. eksport; kroppens loft hæves kun for ruten (16 MB).
   - Audit-logget får `kind: "export"` med mappen som host – kun metadata, aldrig filens indhold.
5. **Fortryd rører ikke filen.** Eksporten sker efter ændringerne i planen, så filen viser arket med de udførte handlinger. *Fortryd* gendanner arbejdsbogen, men filen forbliver – den er et nyt artefakt, som brugeren har bekræftet navnet på.

**Test:** `shared` tester prompt-gating og `checkPlan` (no-bridge, ugyldige navne, forkert vært), og bridgen tester `exportPath` (traversal), `buildXlsx` (round-trip læst tilbage med exceljs) og ruten (200/409/400).
