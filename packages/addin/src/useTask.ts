import { useCallback, useRef, useState } from "react";
import type { TaskRequest, TaskWarning } from "@miai/shared";
import { useBackend } from "./backend";

export interface TaskState {
  status: "idle" | "running" | "done" | "error";
  text: string;
  result?: unknown;
  model?: string;
  progressChars: number;
  /** Værende web-søgning (vises i UI, ADR-0010). */
  searchQuery?: string;
  error?: string;
  /** Fejlkode og HTTP-status fra motoren – oversættes i UI'et (#3). */
  errorCode?: string;
  errorStatus?: number;
  /** Sat, hvis motoren afbrød svaret (fx gentagelsessløjfe). */
  stopReason?: "repetition";
  /** Fx at modellen så mindre, end der blev sendt (kontekstgrænse). */
  warnings?: TaskWarning[];
  /** Fremskridt ved opdeling af lange tekster (map-reduce). */
  step?: { done: number; total: number };
}

const initial: TaskState = { status: "idle", text: "", progressChars: 0 };

/** Kører en opgave mod bridgen og samler den streamede tilstand. */
export function useTask() {
  const [state, setState] = useState<TaskState>(initial);
  const abortRef = useRef<AbortController | null>(null);
  const backend = useBackend();

  const start = useCallback(
    async (id: string, req: TaskRequest, onDelta?: (text: string) => void) => {
      abortRef.current?.abort();
      const abort = new AbortController();
      abortRef.current = abort;
      setState({ ...initial, status: "running" });
      let finalState: TaskState = { ...initial, status: "running" };
      const update = (fn: (s: TaskState) => TaskState) => {
        finalState = fn(finalState);
        setState(finalState);
      };
      try {
        await backend.runTask(
          id,
          req,
          (e) => {
            switch (e.type) {
              case "meta":
                update((s) => ({ ...s, model: e.model }));
                break;
              case "delta":
                onDelta?.(e.text);
                // Når svaret begynder igen efter en søgning, er søgningen afsluttet.
                update((s) => ({ ...s, text: s.text + e.text, searchQuery: e.text ? undefined : s.searchQuery }));
                break;
              case "progress":
                update((s) => ({ ...s, progressChars: e.chars, step: e.step ?? s.step }));
                break;
              case "search":
                update((s) => ({ ...s, searchQuery: e.query }));
                break;
              case "result":
                update((s) => ({ ...s, result: e.data }));
                break;
              case "error":
                update((s) => ({ ...s, status: "error", error: e.message, errorCode: e.code, errorStatus: e.status }));
                break;
              case "retract":
                // Tekst fra en runde, der endte med en søgning, hører ikke til svaret.
                update((s) => ({ ...s, text: s.text.slice(0, Math.max(0, s.text.length - e.chars)) }));
                break;
              case "done":
                update((s) => ({ ...s, status: "done", stopReason: e.stopReason, warnings: e.warnings, step: undefined }));
                break;
            }
          },
          abort.signal,
        );
        if (finalState.status === "running") update((s) => ({ ...s, status: "done" }));
      } catch (err) {
        if (abort.signal.aborted) update((s) => ({ ...s, status: s.text || s.result ? "done" : "idle" }));
        else update((s) => ({ ...s, status: "error", error: (err as Error).message }));
      }
      return finalState;
    },
    [backend],
  );

  const stop = useCallback(() => abortRef.current?.abort(), []);
  const reset = useCallback(() => setState(initial), []);

  return { state, start, stop, reset };
}
