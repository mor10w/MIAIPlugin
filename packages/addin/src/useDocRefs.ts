import { useMemo } from "react";
import type { DocContext } from "@miai/shared";
import type { DocRefs } from "./components/Markdown";
import type { HostAdapter } from "./hosts";
import { findRefs, usedRangeOf } from "./refs";

/**
 * Klikbare henvisninger i svar (#22) for den aktuelle vært: celler og områder i Excel, slides i
 * PowerPoint. Værter uden navigation får ingen. Fejl ved navigation (fx et ark, der ikke findes)
 * ignoreres – et klik må aldrig give en fejlbesked midt i et svar.
 */
export function useDocRefs(adapter: HostAdapter, ctx: DocContext | null | undefined): DocRefs | undefined {
  const used = adapter.host === "excel" ? usedRangeOf(ctx) : undefined;
  return useMemo(() => {
    if (!adapter.goTo || (adapter.host !== "excel" && adapter.host !== "powerpoint")) return undefined;
    return {
      find: (text: string) => findRefs(text, adapter.host, { used }),
      onClick: (ref) => void adapter.goTo!(ref, ctx ?? undefined).catch(() => {}),
    };
    // ctx bruges kun til at finde det aktive ark ved klik; `used` dækker det, der ændrer henvisningerne.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [adapter, used]);
}
