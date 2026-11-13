import { createRoot } from "react-dom/client";
import { Policy } from "@miai/shared";
import { applyCsp, buildCsp } from "./csp";
import { BackendContext, createBackend } from "./backend";
import { App } from "./components/App";
import { createAdapter } from "./hosts";

void Office.onReady(async (info) => {
  const root = createRoot(document.getElementById("root")!);
  const adapter = createAdapter(info.host);
  if (!adapter) {
    root.render(<p style={{ padding: 16 }}>Denne Office-app understøttes ikke endnu. / This Office app is not supported yet.</p>);
    return;
  }
  const backend = await createBackend();
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
