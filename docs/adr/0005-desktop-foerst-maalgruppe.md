# ADR-0005: Lokale desktop-installationer og privat/SMB først

- Status: Vedtaget
- Dato: 2026-11-02

## Kontekst
Første målgruppe er privatpersoner og mindre virksomheder med Office installeret lokalt. Større organisationer, der ikke vil dele data med 3rd party cloud, er en sandsynlig senere målgruppe.

## Beslutning
- Understøt Word, Excel og PowerPoint på Windows og Mac desktop. Office på web er ikke i scope for fase 1–2.
- Add-in'ets filer hostes statisk med direkte forbindelse til modellen, eller serveres af den valgfri bridge (se ADR-0007).
- XML add-in-manifest, sideloadet af installeren (Windows: delt katalog/registry; Mac: `wef`-mappen).
- Enterprise forberedes kun via designregler (politik, konfigurerbar endpoint, audit-log), ikke via funktioner.

## Konsekvenser
+ Ingen hosting-infrastruktur.
+ Enkel sikkerhedsmodel (loopback, same origin).
- Office på web kræver senere, at add-in'et hostes fra en organisations-HTTPS-server; bridgens statiske serving skal derfor holdes adskilt fra API'et.
