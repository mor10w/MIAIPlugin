import { createRoot } from "react-dom/client";
import { Policy } from "@miai/shared";
import { applyCsp, buildCsp } from "./csp";
import { BackendContext, createBackend, type Backend } from "./backend";
import { registerCellFunctions } from "./cell/functions";
import { App } from "./components/App";
import { createAdapter } from "./hosts";

// Regnearksfunktionerne (#80) knyttes til, så snart siden er indlæst: Excel kan beregne celler,
// før task pane er åbnet. Kaldene venter på backend'en. Uden for Excels shared runtime sker intet.
let backendReady!: (b: Backend) => void;
const backendPromise = new Promise<Backend>((ok) => (backendReady = ok));
registerCellFunctions(backendPromise);

void Office.onReady(async (info) => {
  const root = createRoot(document.getElementById("root")!);
  const adapter = createAdapter(info.host);
  if (!adapter) {
    root.render(<p style={{ padding: 16 }}>Denne Office-app understøttes ikke endnu. / This Office app is not supported yet.</p>);
    return;
  }
  const backend = await createBackend();
  backendReady(backend);
  // Office.js kan have defineret CustomFunctions først nu.
  if (adapter.host === "excel") registerCellFunctions(backendPromise);
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
