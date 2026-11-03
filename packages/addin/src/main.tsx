import { createRoot } from "react-dom/client";
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
  root.render(
    <BackendContext.Provider value={backend}>
      <App adapter={adapter} />
    </BackendContext.Provider>,
  );
});
