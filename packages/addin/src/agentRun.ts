import { MAX_READ_ROUNDS, readRequests, type AgentPlan, type DocContext, type TaskRequest } from "@miai/shared";
import type { HostAdapter } from "./hosts";
import type { TaskState } from "./useTask";

type Start = (id: string, req: TaskRequest) => Promise<TaskState>;

/**
 * Agent i flere trin (#20): kører agent-opgaven og lader modellen læse i dokumentet, før den laver
 * sin plan. Returnerer planen med `readRange`, henter add-in'et cellerne (ingen ændringer, ingen
 * bekræftelse – det er samme slags data som markeringen) og spørger igen med værdierne i
 * konteksten. Efter MAX_READ_ROUNDS runder får modellen ikke flere læsninger og skal lave planen.
 *
 * Alle skrivende handlinger vises stadig for brugeren og udføres først efter godkendelse.
 */
export async function planWithReads(
  start: Start,
  adapter: HostAdapter,
  req: { context: DocContext; params: Record<string, string>; history?: TaskRequest["history"] },
  onRead?: (addresses: string[] | null) => void,
): Promise<TaskState> {
  let observations = "";
  try {
    for (let round = 0; ; round++) {
      const canRead = !!adapter.readForAgent && round < MAX_READ_ROUNDS;
      const context = observations ? { ...req.context, observations: observations.slice(0, 40_000) } : req.context;
      const final = await start("agent", { context, params: { ...req.params, agentReads: canRead ? "on" : "" }, history: req.history });
      const plan = final.result as AgentPlan | undefined;
      if (final.status !== "done" || !plan) return final;
      const reads = readRequests(plan, req.context);
      if (!plan.actions.some((a) => a.action === "readRange")) return final;
      // Sidste runde, eller læsninger, der ikke kan udføres: vis resten af planen uden dem.
      if (!canRead || !reads.length) return { ...final, result: { ...plan, actions: plan.actions.filter((a) => a.action !== "readRange") } };
      onRead?.(reads.map((r) => r.address));
      const text = await adapter.readForAgent!(reads, req.context).catch(
        (e: unknown) => `Could not read the cells: ${e instanceof Error ? e.message : String(e)}`,
      );
      observations += (observations ? "\n\n" : "") + text;
    }
  } finally {
    onRead?.(null);
  }
}
