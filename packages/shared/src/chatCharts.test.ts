import { describe, expect, it } from "vitest";
import { AgentPlan, agentToolGuide } from "./agent.js";
import { getTask } from "./tasks/catalog.js";
import { TaskRequest } from "./types.js";

/** Grafer i chatten: i Excel et rigtigt diagram gennem agenten; i de andre apps en Mermaid-figur med klare regler. */

const system = (host: string) =>
  getTask("chat")!.buildMessages(TaskRequest.parse({ context: { host, text: "Aktie\tKurs\nNOVO-B\t274,3" }, params: { prompt: "Dan en trendgraf" } }))[0]
    .content as string;

describe("chattens vejledning om grafer", () => {
  it("Excel: en graf af tal skal være et rigtigt Excel-diagram gennem agenten, med tallene som tabel", () => {
    const sys = system("excel");
    expect(sys).toContain("must be a real Excel chart, never a drawing");
    // Svaret skal nævne agenten: det er dét, der giver brugeren knappen under svaret.
    expect(sys).toContain("MIAI's agent can insert it as an Excel chart");
    expect(sys).toContain("exactly ONE Markdown table for all series");
    // En lokal model tænker ellers højt i svaret, og agenten får kun de første ca. 3.900 tegn som materiale.
    expect(sys).toContain("Never show your calculations");
    // Strukturdiagrammer tegnes stadig.
    expect(sys).toContain("Diagrams of structure (flows, hierarchies, relations) are still drawn with mermaid");
  });

  it("Word og PowerPoint: ingen henvisning til Excel-diagrammer", () => {
    for (const host of ["word", "powerpoint"]) expect(system(host)).not.toContain("real Excel chart");
  });

  it("Excel bliver ikke bedt om at tegne grafer: figurreglen gælder kun strukturdiagrammer, og xychart nævnes ikke", () => {
    const sys = system("excel");
    expect(sys).toContain("a diagram of structure (a flow, hierarchy, process or relations – not a chart of numbers)");
    expect(sys).not.toContain("figure (diagram, chart, grafik)");
    expect(sys).not.toContain("xychart-beta");
  });

  it("Word og PowerPoint: reglerne for Mermaids diagram over tal (kun tal, etiketter i anførselstegn)", () => {
    for (const host of ["word", "powerpoint"]) {
      const sys = system(host);
      expect(sys).toContain("figure (diagram, chart, grafik)");
      expect(sys).toContain("xychart-beta");
      expect(sys).toContain("never null");
      expect(sys).toContain("double-quoted labels without line breaks");
    }
  });

  it("agenten i Excel ved, at tal fra chatten først skrives i arket og derefter får et diagram", () => {
    const guide = agentToolGuide("excel");
    expect(guide).toContain("first write them with setValues");
    expect(guide).toContain("then addChart on exactly that range");
  });

  it("en handling uden forklaring (reason) er stadig en gyldig plan", () => {
    const plan = AgentPlan.parse({
      summary: "Skriver kurserne og laver et diagram.",
      actions: [
        { action: "setValues", address: "D1", values: [["Dato", "Kurs"]], reason: "Kurserne" },
        { action: "addChart", source: "D1:E2", type: "line" },
      ],
    });
    expect(plan.actions[1]).toMatchObject({ action: "addChart", reason: "" });
  });
});
