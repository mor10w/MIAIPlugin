# Sikkerhed

## Rapportér en sårbarhed

Opret **ikke** et offentligt issue for en sikkerhedsfejl. Brug i stedet GitHubs private rapportering: fanen **Security** i repoet → **Report a vulnerability**. Så ser kun vedligeholderne rapporten, indtil fejlen er rettet.

Fortæl gerne:

- hvilken del det drejer sig om (add-in, bridge, MIAI App, workflows) og hvilken udgave
- hvordan fejlen kan genskabes, og hvad en angriber kan opnå
- om den kræver adgang til brugerens maskine, et ondsindet dokument, en ondsindet model eller en ondsindet webside

Du får svar så hurtigt som muligt. Fejlen rettes i en ny udgave, og du kan blive nævnt i udgivelsesnoterne, hvis du ønsker det.

## Understøttede udgaver

Kun den seneste udgave under [Releases](https://github.com/mor10w/MIAIPlugin/releases/latest) får sikkerhedsrettelser. MIAI App opdaterer sig selv på Windows; på Mac skal den nye udgave hentes (se [brugervejledningen](docs/BRUG.md)).

## Hvad der især er relevant

MIAI behandler brugerens dokumenter og sender dem til en sprogmodel. Fejl i disse grænser er sikkerhedsfejl:

- **Politikken for tilladte hosts:** at få add-in'et, bridgen eller websøgningen til at sende indhold til en host, der ikke er tilladt ([ADR-0006](docs/adr/0006-politik-allowed-hosts.md)), eller at omgå en låst organisationspolitik.
- **Bridgen på `localhost`:** at kalde den fra en fremmed webside eller en anden bruger på samme maskine uden adgangstoken ([ADR-0002](docs/adr/0002-lokal-bridge-service.md)).
- **Prompt-injection med virkning:** at et dokument, en mail, en kilde fra brugerens mapper eller et søgeresultat får agenten til at ændre noget, brugeren ikke har godkendt ([ADR-0012](docs/adr/0012-agent-tilstand.md), [ADR-0017](docs/adr/0017-agent-kodetilstand.md)).
- **MIAI App:** certifikatet, der betros lokalt, opdateringsfeedet og installation i Office ([ADR-0016](docs/adr/0016-distribution-signering-opdatering.md)).

Kendte begrænsninger, som ikke er fejl i sig selv: installerne er endnu ikke signeret med et udgivercertifikat (#67), og en model, brugeren selv har valgt, kan give forkerte svar.
