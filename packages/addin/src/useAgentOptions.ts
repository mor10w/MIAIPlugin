import { useEffect, useState } from "react";
import { agentExportAllowed, agentScriptsAllowed, type ExportSheetRequest } from "@miai/shared";
import { useBackend } from "./backend";

/**
 * Agentens tilvalg, som de gælder lige nu: kodetilstand (ADR-0017) og eksport af et ark som .xlsx
 * (tillæg ADR-0012). Begge skal være slået til af brugeren og tilladt af politikken; eksporten
 * kræver desuden bridgen, som er den eneste, der kan skrive filer. `refreshKey` genindlæser dem
 * (fx når brugeren vælger en anden opgave eller sender en ny besked).
 */
export function useAgentOptions(refreshKey?: unknown): { scripts: boolean; exportSheet?: (data: ExportSheetRequest) => Promise<string> } {
  const api = useBackend();
  const [scripts, setScripts] = useState(false);
  const [exportSheet, setExportSheet] = useState<((data: ExportSheetRequest) => Promise<string>) | undefined>();
  useEffect(() => {
    api.settings().then(
      (v) => {
        setScripts(agentScriptsAllowed(v.settings, v.policy));
        setExportSheet(
          api.mode === "bridge" && agentExportAllowed(v.settings, v.policy)
            ? () =>
                (d: ExportSheetRequest): Promise<string> =>
                  api.exportSheet!(d).then((r) => r.path)
            : undefined,
        );
      },
      () => {
        setScripts(false);
        setExportSheet(undefined);
      },
    );
  }, [api, refreshKey]);
  return { scripts, exportSheet };
}
