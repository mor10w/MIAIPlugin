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
- Flyt sker som tekst. Formatering af det flyttede bevares ikke, hvilket er i tråd med "ingen fri formatering".
