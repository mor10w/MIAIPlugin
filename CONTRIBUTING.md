# Bidrag til MIAI

Tak, fordi du vil hjælpe. Projektets sprog er dansk: issues, commit-beskeder, kommentarer i koden og dokumentation skrives på dansk.

## Fejl og forslag

Opret et [issue](https://github.com/mor10w/MIAIPlugin/issues). En god fejlrapport fortæller:

- Hvilken Office-app og version, og om det er Windows eller Mac.
- Hvordan MIAI er installeret (MIAI App, hostet add-in eller fra kildekoden) og hvilken udgave.
- Hvilket inference-værktøj og hvilken model.
- Hvad du gjorde, hvad du forventede, og hvad der skete. Et skærmbillede hjælper.

Slet fortrolig tekst fra eksempler og skærmbilleder, før du deler dem.

## Kom i gang med koden

Du skal bruge Node.js 22 LTS.

```bash
npm install
```

```bash
npm run build
```

```bash
npm test
```

[docs/UDVIKLING.md](docs/UDVIKLING.md) beskriver udviklingsserveren, bridgen, testene og hvordan add-in'et lægges ind i Office. [docs/ARKITEKTUR.md](docs/ARKITEKTUR.md) og [ADR'erne](docs/adr/) forklarer, hvorfor tingene er bygget, som de er.

## Arbejdsgang

- **Ingen direkte commits på `main`.** Lav en gren og en pull request.
  - Gren-navn: `<type>/<kort-beskrivelse>`, fx `fix/mermaid-etiketter` eller `feat/excel-pivot`.
  - Én pull request pr. sammenhængende ændring. Beskriv kort hvad, hvorfor og hvordan det er testet.
- **Merge først, når CI er grøn** (lint, e2e, build på Ubuntu og Windows).
- **Ændrer du adfærd, så opdatér dokumentationen i samme pull request**: [brugervejledningen](docs/BRUG.md) for det, brugeren ser, og en ADR for beslutninger, der er svære at lave om.

## Før du pusher

```bash
npx prettier --check .
```

```bash
npm run lint -- --max-warnings 0
```

```bash
npm test
```

Rører ændringen task pane, så kør også de berørte browsertests:

```bash
npm run build
```

```bash
npx playwright test -c packages/addin/e2e/playwright.config.ts <fil>
```

De visuelle tests (`visual.spec.ts`) kan kun sammenlignes i CI, fordi skrifttyperne er forskellige fra maskine til maskine. Ændrer du brugerfladen med vilje, så opdatér referencebillederne med workflowet *Visuelle baselines* på din gren (Actions → Visuelle baselines → Run workflow).

## Afprøvning i rigtig Office

De automatiske tests kører mod en simuleret udgave af Office. Det fanger meget, men ikke alt, så skriv i pull requesten, om ændringen er prøvet i rigtig Office, og i hvilken app og på hvilken platform. Tjeklisterne står i [docs/UDVIKLING.md](docs/UDVIKLING.md) og [docs/TESTVEJLEDNING.md](docs/TESTVEJLEDNING.md).

## Udgivelser

Du skal ikke selv ændre versionsnummeret. Udgivelser laves automatisk fra `main`; se afsnittet om udgivelse i [docs/UDVIKLING.md](docs/UDVIKLING.md).

## Licens

Koden er udgivet under [MIT-licensen](LICENSE), og ophavsretten til den eksisterende kode tilhører **ManInfo ApS**. Når du sender en pull request, accepterer du, at dit bidrag udgives under samme licens.
