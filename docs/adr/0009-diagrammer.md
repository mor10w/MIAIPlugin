# ADR-0009: Diagrammer – struktureret output med visuelle præstationer i Word og PowerPoint

- Status: Vedtaget
- Dato: 2026-11-02
- Relateret: [ADR-0007](0007-bridge-valgfri-direkte-forbindelse.md), [ADR-0008](0008-distribution-af-solgte-loesninger.md)

## Kontekst
Modellen kan sagtens producere diagrammer (fx Mermaid-rendering virker i andre sammenhænge), men MIAI's opgaver returnerer kun tekst, som indsættes som markdown. Brugeren oplever "flade" resultater: i stedet for en figur får de anvisninger om, hvordan de selv laver SmartArt.

Research (Microsoft Learn, 2026-11) viser at Office.js understøtter det nødvendige:

| Host | Mulighed | API |
|---|---|---|
| Word | Indsæt billede ved markering (inline eller svævende) | `Range.insertInlinePictureFromBase64(base64, "After")`, `insertPictureFromBase64` |
| PowerPoint | **Native figurer** – node pr. shape, navigerbare/redigerbare af brugeren | `slide.shapes.addGeometricShape(type, {left,top,width,height})` + `shape.textFrame.textRange.text`; linjer: `shapes.addLine(ConnectorType.straight, …)` (PowerPointApi 1.4); gruppering: `shapes.addGroup` (1.8) |

Begrænsninger i PowerPoint-API'et (verificeret i `@types/office-js`, 2026-11-03):
- `addLine` tager kun en afgrænsningsboks. Negative `width`/`height` giver `InvalidArgument`, og der findes ingen flip. Skrå linjer kan derfor kun gå fra øverst-venstre mod nederst-højre.
- `ShapeLineFormat` har ingen pilehoveder.
- `GeometricShapeType.flowChartConnector` er flowchart-symbolet "forbindelsespunkt" (en cirkel), ikke en linje.
| PowerPoint | Billede som shape-fill (alternativ) | `shape.fill.setImage(base64)` (PowerPointApi 1.8) |

Mermaid er velegnet til *visning* (GitHub, VS Code), men leverer ikke det objektmodel, der skal til for native shapes. Dermed er Mermaid en renderer, ikke et format: vi definerer et simpelt JSON-skema (knuder + kanter), som modellen nemt producerer og add-in'et selv renderer.

## Beslutning
Ny opgave **"Lav diagram"** (`diagram`), der konverterer den markerede tekst/dokument til et struktureret diagram:

1. **Format:** `DiagramResult = { title?, nodes: [{id, label}], edges: [{from, to}] }` (zod-skema i `packages/shared/src/diagram.ts`). Kortere og mere pålideligt for lokale modeller end Mermaid-syntaks; ingen parser afhængighed.
2. **Layout og ruter** (rene funktioner i `shared`, testet):
   - `layoutDiagram()` bryder først cykler: en DFS finder tilbagekanter, så en feedback-løkke ikke lægger hele diagrammet i én søjle. Derefter lagdeles efter længste sti, venstre mod højre.
   - `routeEdges()` laver **vinkelrette** ruter, hvor alle segmenter er vandrette eller lodrette. Fremadkanter knækker i mellemrummet foran målets søjle. Returkanter løber i egne baner under diagrammet. Alle kanter ender i målets venstre side, og ingen lodrette segmenter krydser en knude.
   - `diagramToSvg()` bruger samme ruter. Forhåndsvisning i task pane, Word-billede og PowerPoint-figurer ser derfor ens ud.
3. **Indsætning:**
   - *Word:* SVG → PNG (2x) via canvas. Billedet indsættes i et nyt afsnit efter markeringen (eller sidst i dokumentet) i naturlig størrelse, højst 450 pt bred, med alt-tekst.
   - *PowerPoint:* **native figurer** på den valgte slide. Knuder er afrundede rektangler med tekst. Kanter er lige linjesegmenter langs de vinkelrette ruter, med en lille chevron som pilehoved (den peger mod højre uden rotation). Skriftstørrelsen skaleres med diagrammet. Med PowerPointApi 1.8 grupperes alt, så diagrammet kan flyttes samlet; gruppen kan opløses for at redigere enkeltdele.
4. **Prompt:** 3–10 knuder, labels maks. ~6 ord, JSON-only (samme mønster som `text-to-slides`; oMLX har `structuredOutput: "none"`, så validering sker via `extractJson` + retry i motoren).

## Konsekvenser
+ Virkelige figurer i Word (billede) og PowerPoint (native, redigerbare objekter) – kerneværdi i salgsudgave B.
+ Renderer er deterministisk og testbar; modellen skal kun producere simpelt JSON.
+ Word-fiksen er et billede, ikke SmartArt – kan ikke redigeres som native objekt (Word.js har ingen shape-API); tilstrækkeligt for v1.
+ PowerPoint-layout er automatisk; brugeren kan flytte figurerne bagefter. Linjerne er ikke "forbundet" til figurerne (API'et kan ikke det), så de følger ikke med, hvis en enkelt figur flyttes.
+ Canvas-baseret SVG→PNG kræver systemfont i webview'et; verificeres på Word/PowerPoint-webviews (WebKit).
