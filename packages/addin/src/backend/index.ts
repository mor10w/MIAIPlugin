import { createContext, useContext } from "react";
import { BridgeBackend, detectBridge } from "./bridge";
import { DirectBackend } from "./direct";
import type { Backend } from "./types";

/**
 * Vælger backend automatisk: serveres task pane af en MIAI Bridge, bruges den.
 * Ellers (fx hostet statisk på GitHub Pages) taler add-in'et direkte med modellen.
 */
export async function createBackend(): Promise<Backend> {
  // Hostet add-in: organisationens konfiguration (miai-config.json) hentes fra samme adresse (#7).
  return (await detectBridge()) ? new BridgeBackend() : DirectBackend.create();
}

export const BackendContext = createContext<Backend | null>(null);

export function useBackend(): Backend {
  const b = useContext(BackendContext);
  if (!b) throw new Error("BackendContext mangler");
  return b;
}

export type { Backend, BackendMode } from "./types";
export { corsHint } from "./direct";
