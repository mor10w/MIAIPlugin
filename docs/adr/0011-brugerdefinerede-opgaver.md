# ADR-0011: Brugerdefinerede opgaver som redigerbare filer (F14)

- **Status:** Beslut
- **Dato:** 2026-11-03

## Kontekst

Opgavekataloget er i dag indbygget i koden (`shared/tasks.ts`). Brugere og organisationer
vil definere egne opgaver (fx "Skriv tilbudstekst", "Opsummér i kundens tone") uden at
vente på en opdatering. PLAN.md fase 2 kræver opgaver som *brugerredigerbare filer, der
deles via mappe* – i en organisation skal IT kunne uddelte opgaver til alle brugere,
uden at de kan ændres.

## Beslutning

1. **Format:** En opgave er én JSON-fil, valideret af `UserTaskFile` (Zod) i `shared`.
   Filen deklareres: id, lokaliserede navn/beskrivelse (da valgfrit en), apps
   (word/excel/powerpoint), input (selection / selectionOrDocument / prompt), output-
   format (plain / markdown / tsv – ikke formula/JSON, som kræver validering i koden),
   parametre (maks. 8), anvendelsesmuligheder, temperatur og en prompt-skabelon
   (system + bruger). Skabelonen udfyldes med `{{placeholderer}}`: `selection`,
   `prompt`, `language`, `address`, `table` og opgavens parametre. Ukendte
   placeholderer efterlates ubehandlede, så brugeren opdager dem. Omsættes til en
   `TaskDefinition` via `toTaskDefinition()`, så den fælles motor kører fil-opgaver
   identisk med de indbyggede.
2. **Prompt-injection:** Hvis skabelonen bruger `{{selection}}`, pakkes markeringen i
   `<document>`-tags – samme sikring som de indbyggede opgaver.
3. **Bridgen (mappe-baseret):** Brugers opgaver læses fra `settings.customTasksDir`
   (standard `<config>/tasks`), organisationens fra `MIAI_TASKS_DIR`. Maskin-opgaver
   vinder ved id-kollision og er read-only via UI (409). En bruger-opgave *må*
   oversteje en indbygget id (organisationstilpasset prompt). Mappen læses med 5 s
   cache, invalideres ved skrift. API: `GET/PUT/DELETE /api/user-tasks[/:id]`,
   `GET /api/tasks` returnerer det samlede katalog.
4. **Direkte tilstand (browser):** Browsere kan ikke læse mapper, så bruger-opgaver
   gemmes i localStorage (importér/eksportér af én fil ad gangen via "Opgaver"-fanen).
   Kollision med en indbygget id afvises, da der ingen maskin-mappe er at
   differentiere fra.

## Konsekvenser

- **Fordele:** Nye opgaver kræver ingen kode eller geninstallation; deling er
  "kopier mappen/filen". Organisationer kan uddelte standardopgaver, og brugeren kan
  tilpasse dem lokalt.
- **Begrænsninger:** Filer kan kun udtrykke tekstoutput (plain/markdown/tsv);
  strukturerede opgaver (grammatik, formler, diagrammer) kræver stadig kode.
  Direkte tilstand kan ikke dele opgaver via mappe – kun fil-import/eksport.
- **Sikkerhed:** Maskin-opgaver kan ikke slettes/redigeres via UI; indhold i filer
  er brugers eget ansvar (samme status som en prompt skrevet i chat).
