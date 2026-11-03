# ADR-0014: Præsentationsdesign – master og layouts ud fra indholdet

- **Status:** Besluttet
- **Dato:** 2026-11-04

## Kontekst

I PowerPoint skal MIAI kunne lave en designskabelon (slidemaster) til hele præsentationen ud fra indholdet og lægge præsentationen ud, så alle slides følger samme design. Dernæst skal agenten (ADR-0012) kunne bygge videre på de samme layouts.

PowerPoints JavaScript-API kan læse mastere og layouts, men **ikke oprette eller ændre dem**. Det kan derimod indsætte slides fra en .pptx-fil (`insertSlidesFromBase64`, PowerPointApi 1.2). Med *behold kildeformatering* følger filens master og layouts med ind i præsentationen.

## Beslutning

1. **Modellen vælger, skabelonerne tegner.**
   - Opgaven `design-deck` giver modellen præsentationens slides (eller et emne, hvis den er tom) og designregler.
   - Modellen svarer med en `DesignPlan`:
     - navn og begrundelse
     - stil: `clean`, `bold` eller `editorial`
     - palet med fem roller
     - skrifttyper fra en liste over skrifttyper, der følger med Office
     - indholdet lagt ud på syv layouttyper: titel, sektion, punkter, to kolonner, citat, nøgletal og afslutning
   - Geometrien pr. stil og layout ligger i faste skabeloner (`addin/src/deck/templates.ts`). Modellen kan altså ikke lave et rodet layout, og designet bliver ens gennem præsentationen.
2. **Farverne tjekkes, før de bruges** (`fixPalette`):
   - Tekst og sekundær tekst skal have ≥ 4,5:1 mod baggrunden, og accent ≥ 3:1.
   - Ellers justeres farven i små trin mod mørk eller lys.
   - Tekst på primærfladen er hvid eller næsten sort efter bedste kontrast.
3. **Bygning:** add-in'et laver en .pptx med pptxgenjs:
   - Ét layout pr. brugt slidetype, altid inklusive *Punkter*, så agenten kan lave indholdsslides. Layoutene får navne som "Nordisk klar – Punkter" og har titel- og tekstpladsholdere med designets skrifttyper og farver.
   - Temaets skrifttyper sættes også.
   - Slides med indholdet i pladsholderne.
   - Nøgletal står som tekstbokse placeret efter antal.
4. **Indsættelse:** filen indsættes efter de nuværende slides med kildeformatering.
   - Vælger brugeren *Erstat de nuværende slides*, gemmes de gamle først (`exportAsBase64`, PowerPointApi 1.8) og fjernes bagefter.
   - *Fortryd* fjerner designets slides og lægger de gamle tilbage i rækkefølge med al formatering.
   - Uden 1.8 bevares de gamle slides, og brugeren får besked.
5. **Forhåndsvisning før indsættelse:**
   - Task pane viser navn, begrundelse, palet (med hex), skrifttyper og en SVG-miniature af hvert slide.
   - Miniaturerne tegnes med de samme skabeloner som .pptx-filen, så det, man ser, er det, man får.
   - Intet ændres, før brugeren trykker *Opret*.

## Konsekvenser

- Designet anvendes ved at bygge præsentationen op i nye slides, ikke ved at ændre de eksisterende, for det kan API'et ikke.
  - Tekstindholdet flyttes med.
  - Billeder, diagrammer, noter og animationer på de gamle slides flyttes ikke. Derfor er *erstat* fravalgt som standard, og de gamle kan altid komme tilbage med fortryd.
- Masteren bliver i præsentationen efter fortryd: PowerPoint fjerner ikke ubrugte mastere via API'et. Den kan slettes i Masterversion.
- Layoutene kan bruges fra *Nyt slide* i PowerPoint og af agentens `addSlide`, så nye slides følger designet.
- Skrifttyper uden for Office' standardsæt er ikke tilladt. Aptos kræver Microsoft 365, og ellers erstatter PowerPoint den pænt.
- Afhængighed: pptxgenjs (MIT). Den indlæses først, når designet oprettes (dynamisk import), så task pane ikke bliver tungere at åbne.

**Test:**
- Unit-tests pakker den genererede .pptx ud og tjekker master, layoutnavne, pladsholdere, skrifttyper, farver og antal slides.
- Skabelonerne holder sig inden for slidets mål for alle stilarter og layouts, og paletten rettes for kontrast.
- E2e tjekker forhåndsvisning (inklusive axe), indsættelse med kildeformatering, erstatning og fortryd, og opførsel i ældre PowerPoint.
- Om master og layouts ser rigtige ud, skal bekræftes i rigtig PowerPoint; tjeklisten står i UDVIKLING.md.
