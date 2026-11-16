import { describe, expect, it } from "vitest";
import { chosenModel, MODEL_GROUPS, modelGroupOf } from "./modelGroups.js";
import { TASKS } from "./tasks.js";
import { Settings } from "./types.js";

describe("model pr. opgavetype", () => {
  it("hver indbygget opgave hører til præcis én type – også nye opgaver skal placeres", () => {
    for (const task of TASKS) {
      const groups = MODEL_GROUPS.filter((g) => (g.tasks as readonly string[]).includes(task.id));
      expect(
        groups.map((g) => g.id),
        task.id,
      ).toHaveLength(1);
    }
    // Ingen type nævner en opgave, der ikke findes.
    const ids = new Set(TASKS.map((t) => t.id));
    for (const g of MODEL_GROUPS) for (const id of g.tasks) expect(ids.has(id), `${g.id}: ${id}`).toBe(true);
  });

  it("3–5 typer, så valget er overskueligt", () => {
    expect(MODEL_GROUPS.length).toBeGreaterThanOrEqual(3);
    expect(MODEL_GROUPS.length).toBeLessThanOrEqual(5);
  });

  it("egne og ukendte opgaver hører under Ret og omskriv", () => {
    expect(modelGroupOf("min-egen-opgave")).toBe("edit");
    expect(modelGroupOf("agent")).toBe("agent");
  });

  it("opgavens eget valg, så typens, så standardmodellen", () => {
    const s = (extra: object) =>
      Settings.parse({ endpoints: [{ id: "x", name: "X", baseUrl: "http://127.0.0.1:1/v1" }], activeEndpoint: "x", defaultModel: "standard", ...extra });
    const settings = s({ groupModels: { edit: "lille", agent: "stor" }, taskModels: { translate: "oversætter" } });
    expect(chosenModel(settings, "simplify")).toBe("lille");
    expect(chosenModel(settings, "translate")).toBe("oversætter");
    expect(chosenModel(settings, "agent")).toBe("stor");
    expect(chosenModel(settings, "chat")).toBe("standard");
    expect(chosenModel(s({ defaultModel: "" }), "chat")).toBe("");
  });
});
