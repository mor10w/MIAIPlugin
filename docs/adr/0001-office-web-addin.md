# ADR-0001: Office Web Add-in (Office.js) frem for VSTO/COM

- Status: Foreslået
- Dato: 2026-11-02

## Kontekst
Produktet skal virke i Word, Excel, PowerPoint og senere Outlook. VSTO/COM giver fuld objektmodel, men kun på Windows-desktop, og Microsoft prioriterer web add-ins.

## Beslutning
Byg som Office Web Add-in (task pane + ribbon-kommandoer) med Office.js, TypeScript og React.

## Konsekvenser
+ Windows, Mac og Office på web; central udrulning via M365 Admin Center.
+ Én kodebase for alle hosts.
- Begrænset af Office.js requirement sets (især PowerPoint).
- Kræver HTTPS og løsning for adgang til localhost (se ADR-0002).
