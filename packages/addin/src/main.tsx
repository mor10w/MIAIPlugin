import { createRoot } from "react-dom/client";
import { Policy } from "@miai/shared";
import { applyCsp, buildCsp } from "./csp";
import { BackendContext, createBackend } from "./backend";
import { registerCellFunctionsWhenReady } from "./cell/functions";
import { App } from "./components/App";
import { createAdapter } from "./hosts";

// Regnearksfunktionerne (#80) må ikke afhænge af panelets opstart. Excel kan køre dem i sin egen,
// skjulte udgave af siden, hvor Office.onReady ikke nødvendigvis melder Excel som vært – og venter
// funktionen på panelet dér, står cellen på #BUSY! for evigt. Backend'en oprettes derfor med det
// samme (den bruger kun fetch og sidens adresse), og funktionerne knyttes til, så snart Office.js
// har defineret CustomFunctions.
const backendPromise = createBackend();
registerCellFunctionsWhenReady(backendPromise);

void Office.onReady(async (info) => {
  const root = createRoot(document.getElementById("root")!);
  const adapter = createAdapter(info.host);
  if (!adapter) {
    root.render(<p style={{ padding: 16 }}>Denne Office-app understøttes ikke endnu. / This Office app is not supported yet.</p>);
    return;
  }
  const backend = await backendPromise;
  // CSP (#28): browseren håndhæver politikkens hosts, før noget af dokumentet kan sendes nogen steder hen.
  // Kan politikken ikke hentes, gælder standarden (kun maskinen selv).
  const policy = await backend.settings().then(
    (v) => v.policy,
    () => Policy.parse({}),
  );
  applyCsp(buildCsp(backend.mode, policy));
  root.render(
    <BackendContext.Provider value={backend}>
      <App adapter={adapter} />
    </BackendContext.Provider>,
  );
});
