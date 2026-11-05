import { describe, expect, it } from "vitest";
import { repairMermaid } from "./mermaid";

describe("repairMermaid", () => {
  it("sætter anførselstegn om etiketter med parenteser og specialtegn (regression: parse error i Word-chatten)", () => {
    expect(repairMermaid("graph TD\n  R2[Platformarkitekt (eksisterende gruppe)]:::role")).toBe(
      'graph TD\n  R2["Platformarkitekt (eksisterende gruppe)"]:::role',
    );
    expect(repairMermaid("flowchart LR\n  A[Start] --> B{Er det ok?} -->|ja| C(Færdig (helt))")).toBe(
      'flowchart LR\n  A["Start"] --> B{"Er det ok?"} -->|ja| C("Færdig (helt)")',
    );
    expect(repairMermaid('graph TD\n  A["Allerede (citeret)"] --> DB[(Database)]')).toBe('graph TD\n  A["Allerede (citeret)"] --> DB[("Database")]');
    expect(repairMermaid("graph TD\n  subgraph P1 [Domæne & (Modellering)]\n  end")).toBe('graph TD\n  subgraph P1 ["Domæne & (Modellering)"]\n  end');
    expect(repairMermaid('graph TD\n  A[Han sagde "hej"]')).toBe('graph TD\n  A["Han sagde #quot;hej#quot;"]');
  });

  it("gør <b>/<i> til markdown-etiketter (de vises ellers som tekst uden htmlLabels)", () => {
    expect(repairMermaid("graph TD\n  A[<b>Dataarkitekt</b><br/><i>(note)</i>]:::header")).toBe('graph TD\n  A["`**Dataarkitekt**\n*(note)*`"]:::header');
    // Kun <br/> kan blive stående i en almindelig etiket.
    expect(repairMermaid("graph TD\n  A[Linje 1<br/>Linje 2]")).toBe('graph TD\n  A["Linje 1<br/>Linje 2"]');
  });

  it("rører ikke andre diagramtyper, stil-linjer eller den nye @{}-syntaks", () => {
    const seq = "sequenceDiagram\n  A->>B: Hej (igen)";
    expect(repairMermaid(seq)).toBe(seq);
    const styles = 'graph TD\n  classDef x fill:#fff,stroke:#000\n  style A fill:#f9f\n  A@{ shape: rect, label: "x" }';
    expect(repairMermaid(styles)).toBe(styles);
  });
});
