# ADR-0015: Outlook og central udrulning (Office på web)

- **Status:** Besluttet
- **Dato:** 2026-11-04
- **Issue:** #7

## Kontekst

Fase 3 (PLAN.md §6) har to krav:
- **Outlook:** MIAI i Outlook til at opsummere, besvare og omskrive mails.
- **Central udrulning:** organisationer udruller add-in'et centralt, også til Office på web, og peger det på en intern inference-server, så data ikke forlader organisationen.

Gennemgangen af #7 præciserede tre ting:
- Office på web kører i browseren. Den kan i princippet nå loopback, men i en organisation peger add-in'et typisk på en intern server. Det kræver et hostet add-in, ikke en bridge.
- I direkte tilstand ligger politikken i `localStorage` og kan ændres af brugeren. Ved central udrulning skal den kunne leveres centralt.
- Outlook bruger et andet API (Mailbox) og et andet manifest (MailApp).

## Beslutning

### Outlook

1. **Samme task pane, egen adapter.** `OutlookAdapter` (`addin/src/hosts/outlook.ts`) håndterer de to tilstande:
   - **Læsetilstand:** mailen med emne og afsender er konteksten (`body.getAsync`). Svar åbnes som udkast med `displayReplyForm`/`displayReplyAllForm`. MIAI sender aldrig selv.
   - **Skrivetilstand:** markeringen, ellers hele udkastet, er konteksten. Resultatet indsættes som HTML eller tekst efter udkastets format: erstat markeringen (`setSelectedDataAsync`), hele udkastet (`body.setAsync`) eller efter markeringen.
2. **Opgaver:**
   - Opgaver, der giver mening for mails, har Outlook som vært: ret, omskriv, parafrasér, forenkl, opsummér, oversæt, salgstekst og chat.
   - Den nye opgave **Svar på mail** tager tone og "hvad svaret skal sige" og må ikke opfinde aftaler. Den indsætter `[pladsholdere]` i stedet.
   - Agent, design og Excel-opgaver vises ikke i Outlook.
3. **Manifest:**
   - `manifest.outlook.template.xml` (MailApp, Mailbox 1.5, ReadWriteItem, knapper i læse- og skrivetilstand) med egne Id'er.
   - `make-manifest.mjs` laver bridge-, dev- og hostede varianter. GitHub Pages publicerer `manifest-outlook.xml`.
   - Bridgen og MIAI App skriver et personligt `manifest-outlook.xml` med port og adgangstoken.
   - Outlook installeres manuelt fra fil, for Outlook har ingen sideload-mappe som Word, Excel og PowerPoint.

### Central udrulning

4. **`miai-config.json` på add-in'ets egen adresse.** Et hostet add-in henter ved start filen ved siden af `taskpane.html`, med formatet `{ policy, settings }`:
   - **Politikken** har forrang, er altid låst og kan ikke ændres i add-in'et. I Indstillinger står der "Organisationen (central udrulning)".
   - **`settings`** er standard for brugere, der ikke selv har gemt noget. Typisk er det organisationens interne endpoint.
   - **Ingen fil (404) eller ingen forbindelse:** add-in'et virker som før.
   - **En fil, der ikke kan læses:** politikken låses med standardens tilladte hosts (fail closed). En ødelagt fil må aldrig åbne for flere hosts.
5. **Ingen bridge til Office på web.** Add-in'et hostes på organisationens egen webserver, og modellen kører på en intern server, der tillader CORS fra add-in'ets adresse. Udrulning sker via M365 Admin Center → Integrated apps med de hostede manifester for Office-apperne og Outlook.

## Konsekvenser

- Alt i MIAI kører fortsat uden en vendor-cloud: add-in'et ligger på organisationens server, og modellen på organisationens inference-server.
- Brugeren kan ikke omgå organisationens politik i add-in'et. Den, der kan ændre `miai-config.json` på serveren, styrer politikken, og det er IT.
- Outlook har ikke sporede ændringer. I skrivetilstand er Outlooks egen fortryd (Ctrl+Z) vejen tilbage, og svar åbnes altid som udkast.
- Grammatik i Outlook erstatter markeringen med den rettede tekst. Formatering inde i markeringen kan gå tabt, som i PowerPoint.
- Outlook på web og nyt Outlook til Windows bruger samme manifest.

**Test:**
- E2e med simuleret Outlook: læsetilstand (opgavevalg, prompt med emne og afsender, svar som udkast med HTML), skrivetilstand (erstat som HTML) og chat om mailen.
- E2e for central konfiguration: låst politik, fail closed og uden fil.
- Bridge-test af det personlige Outlook-manifest.
- Rigtig Outlook og en rigtig central udrulning skal afprøves; tjeklisten står i UDVIKLING.md.
