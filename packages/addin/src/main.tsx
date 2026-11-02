import { createRoot } from "react-dom/client";
import { App } from "./components/App";
import { createAdapter } from "./hosts";

Office.onReady((info) => {
  const root = createRoot(document.getElementById("root")!);
  const adapter = createAdapter(info.host);
  root.render(adapter ? <App adapter={adapter} /> : <p style={{ padding: 16 }}>Denne Office-app understøttes ikke endnu. / This Office app is not supported yet.</p>);
});
