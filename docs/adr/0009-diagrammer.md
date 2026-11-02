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
| PowerPoint | **Native figurer** – node pr. shape, navigerbare/redigerbare af brugeren | `slide.shapes.addGeometricShape(type, {left,top,width,height})` + `shape.textFrame.textRange.text`; linjer: `GeometricShapeType.flowChartConnector` med `flipHorizontal/flipVertical` |
| PowerPoint | Billede som shape-fill (alternativ) | `shape.fill.setImage(base64)` (PowerPointApi 1.8) |

Mermaid er velegnet til *visning* (GitHub, VS Code), men leverer ikke det objektmodel, der skal til for native shapes. Dermed er Mermaid en renderer, ikke et format: vi definerer et simpelt JSON-skema (knuder + kanter), som modellen nemt producerer og add-in'et selv renderer.

## Beslutning
Ny opgave **"Lav diagram"** (`diagram`), der konverterer den markerede tekst/dokument til et struktureret diagram:

1. **Format:** `DiagramResult = { title?, nodes: [{id, label}], edges: [{from, to}] }` (zod-skema i `packages/shared/src/diagram.ts`). Kortere og mere pålideligt for lokale modeller end Mermaid-syntaks; ingen parser afhængighed.
2. **Rendering:** `layoutDiagram()` (rengt-lagdeling, venstre-modet-højre) + `diagramToSvg()` – rene funktioner i shared, testet. Taskpane'et viser et SVG-forhåndsgreb; til Word konverteres SVG→PNG via canvas.
3. **Indsætning:**
   - *Word:* PNG-billede indsættes inline ved markeringen (`insertInlinePictureFromBase64`).
   - *PowerPoint:* **native shapes** – afrounded rektangler med label + `flowChartConnector`-linjer mellem knuderne, placeret på den valgte slide. Brugeren kan derefter klikke på, flytte og redigere hver del (kravet om at "navigere rundt i objekter").
4. **Prompt:** 3–10 knuder, labels maks. ~6 ord, JSON-only (samme mønster som `text-to-slides`; oMLX har `structuredOutput: "none"`, så validering sker via `extractJson` + retry i motoren).

## Konsekvenser
+ Virkelige figurer i Word (billede) og PowerPoint (native, redigerbare objekter) – kerneværdi i salgsudgave B.
+ Renderer er deterministisk og testbar; modellen skal kun producere simpelt JSON.
+ Word-fiksen er et billede, ikke SmartArt – kan ikke redigeres som native objekt (Word.js har ingen shape-API); tilstrækkeligt for v1.
+ PowerPoint-layout er automatisk (ingen manuel justering i add-in'et); brugeren flytter shapes manuelt efter indsættelse.
+ Canvas-baseret SVG→PNG kræver systemfont i webview'et; verificeres på Word/PowerPoint-webviews (WebKit).
