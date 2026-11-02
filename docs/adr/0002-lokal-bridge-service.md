# ADR-0002: Lokal bridge-service mellem add-in og model

- Status: Erstattet af [ADR-0007](0007-bridge-valgfri-direkte-forbindelse.md). Bridgen er nu valgfri.
- Dato: 2026-11-02

## Kontekst
Add-in kører i en webview over HTTPS. Direkte kald til lokale model-runtimes rammer mixed content, CORS og Private Network Access. Cloud-nøgler må ikke ligge i webview, og RAG kræver filsystemadgang.

## Beslutning
En lille lokal service ("MIAI Bridge") på `https://127.0.0.1` med lokalt installeret certifikat og pr.-installation token. Add-in taler kun med bridgen.

## Konsekvenser
+ Ét sted for provider-logik, prompts, logging og sikkerhed.
+ Kan levere add-in'ets statiske filer lokalt (offline).
- Ekstra installationstrin (installer håndterer service + cert).
- Skal verificeres i fase 0-spike på alle platforme.
