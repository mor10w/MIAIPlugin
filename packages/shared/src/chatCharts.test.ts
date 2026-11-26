import { describe, expect, it } from "vitest";
import { AgentPlan, agentToolGuide } from "./agent.js";
import { chartMaterial } from "./chartMaterial.js";
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

describe("chatsvaret som materiale til agenten", () => {
  const INIT = "%%{init: {'theme':'base', 'themeVariables': {'xyChart': {'titleColor': '#333333'}}}}%%";
  const chart = (title: string) =>
    [
      "```mermaid",
      INIT,
      "xychart-beta",
      `    title "${title}"`,
      '    x-axis ["16-sep", "17-sep", "18-sep"]',
      '    y-axis "Kurs (DKK)" 245 --> 290',
      "    line [279.35, 281.55, 260]",
      "    line [267.4, 270.5]",
      "```",
    ].join("\n");

  it("erstatter en graf af tal med dens tal, uden diagramkoden", () => {
    const m = chartMaterial(`Her er grafen.\n\n${chart("NOVO.B - 30 dage")}\n\nSlut.`);
    expect(m).toContain("Chart: NOVO.B - 30 dage");
    expect(m).toContain("x (3): 16-sep, 17-sep, 18-sep");
    expect(m).toContain("y: Kurs (DKK)");
    expect(m).toContain("series 1 (3): 279.35, 281.55, 260");
    expect(m).toContain("series 2 (2): 267.4, 270.5");
    expect(m).toContain("Her er grafen.");
    expect(m).toContain("Slut.");
    expect(m).not.toMatch(/mermaid|xychart|init|```/);
  });

  it("holder tre grafer inden for grænsen for materiale, så ingen aktie mangler", () => {
    const labels = Array.from({ length: 33 }, (_, i) => `"${i + 1}-okt"`).join(", ");
    const values = Array.from({ length: 33 }, (_, i) => (18_000 + i * 12.35).toFixed(2)).join(", ");
    const big = (title: string) =>
      ["```mermaid", INIT, "xychart-beta", `    title "${title}"`, `    x-axis [${labels}]`, `    line [${values}]`, "```"].join("\n");
    const m = chartMaterial(["NOVO.B", "GENMAB", "MAERSK.B"].map(big).join("\n\n"));
    expect(m.length).toBeLessThan(3_900);
    for (const name of ["NOVO.B", "GENMAB", "MAERSK.B"]) expect(m).toContain(`Chart: ${name}`);
  });

  it("læser et cirkeldiagram og en graf, modellen ikke nåede at lukke", () => {
    expect(chartMaterial('```mermaid\npie title Fordeling\n    "A" : 40\n    "B" : 60\n```')).toBe("Chart: Fordeling\nA\t40\nB\t60\n");
    expect(chartMaterial('```mermaid\nxychart-beta\n    x-axis ["a", "b"]\n    bar "Salg" [1, 2]')).toContain('series 1 "Salg" (2): 1, 2');
  });

  it("lader andre figurer og almindelig tekst stå", () => {
    const flow = "Se her:\n\n```mermaid\nflowchart TD\n    A --> B\n```\n\n| a | b |\n|---|---|\n| 1 | 2 |";
    expect(chartMaterial(flow)).toBe(flow);
  });
});

describe("agentens vejledning om tal fra chatten", () => {
  const agent = (params: Record<string, string>) =>
    getTask("agent")!.buildMessages(
      TaskRequest.parse({ context: { host: "excel", text: "Aktie\tKurs" }, params: { prompt: "Dan en trendgraf", ...params } }),
    )[0].content as string;

  it("siger, at agenten ikke kan hente noget på nettet, heller ikke med et script", () => {
    const s = agent({ agentScripts: "on" });
    expect(s).toContain("no access to the internet, and neither has runScript");
    expect(s).toContain("<material> are data you already have");
  });

  it("lader ikke læsning i arket stå i vejen for tal fra chatten", () => {
    expect(agent({ agentReads: "on" })).toContain("never refuse because it is not in the sheet yet");
    expect(agent({})).not.toContain("never refuse because it is not in the sheet yet");
  });
});
