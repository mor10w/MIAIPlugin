import type { HealthInfo, Policy, Settings } from "@miai/shared";

/**
 * Statusside på "/" (#15). Vises, når nogen åbner bridgens adresse i en almindelig browser
 * (fx "Åbn https://localhost:44321" i MIAI App). Selve UI'et (taskpane.html) kører kun i Office.
 * Renderes på serveren uden JavaScript og uden afhængigheder.
 */

export interface StatusInput {
  version: string;
  health: HealthInfo;
  settings: Settings;
  policy: Policy;
  policySource: string;
  addinBuilt: boolean;
  configFile: string;
  lang: "da" | "en";
}

const T = {
  da: {
    title: "MIAI Bridge",
    running: "Bridgen kører",
    version: "Version",
    endpoint: "Aktiv endpoint",
    reachable: "Svarer",
    unreachable: "Svarer ikke",
    endpoints: "Konfigurerede endpoints",
    policy: "Politik",
    policySource: { default: "standard", user: "brugerens", machine: "organisationens (låst)", browser: "add-in'ets" } as Record<string, string>,
    allowedHosts: "Tilladte hosts",
    webSearch: "Web-søgning",
    auditLog: "Audit-log",
    on: "til",
    off: "fra",
    forbidden: "forbudt af politikken",
    config: "Konfiguration",
    addin: "Add-in",
    addinOk: "Bygget og klar. Manifestet peger på",
    addinMissing: "Add-in'et er ikke bygget. Kør npm run build.",
    officeOnly:
      "MIAI's brugerflade kører kun i Word, Excel og PowerPoint (fanen Hjem → AI-assistent). Denne side viser bare, at bridgen kører.",
    api: "API",
    apiNote: "API'et under /api/ kræver headeren X-MIAI og kan derfor ikke åbnes direkte i browseren. Rå status:",
  },
  en: {
    title: "MIAI Bridge",
    running: "The bridge is running",
    version: "Version",
    endpoint: "Active endpoint",
    reachable: "Reachable",
    unreachable: "Not reachable",
    endpoints: "Configured endpoints",
    policy: "Policy",
    policySource: { default: "default", user: "user", machine: "organisation (locked)", browser: "add-in" } as Record<string, string>,
    allowedHosts: "Allowed hosts",
    webSearch: "Web search",
    auditLog: "Audit log",
    on: "on",
    off: "off",
    forbidden: "forbidden by policy",
    config: "Configuration",
    addin: "Add-in",
    addinOk: "Built and ready. The manifest points to",
    addinMissing: "The add-in is not built. Run npm run build.",
    officeOnly: "MIAI's user interface only runs inside Word, Excel and PowerPoint (Home tab → AI assistant). This page just shows that the bridge is running.",
    api: "API",
    apiNote: "The API under /api/ requires the X-MIAI header and cannot be opened directly in a browser. Raw status:",
  },
};

const esc = (s: unknown) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Vælger sprog ud fra Accept-Language (dansk, ellers engelsk). */
export function pickLang(acceptLanguage: string | undefined): "da" | "en" {
  const first = (acceptLanguage ?? "").split(",")[0]?.trim().toLowerCase() ?? "";
  return first.startsWith("da") || first.startsWith("nb") || first.startsWith("nn") ? "da" : "en";
}

export function renderStatusPage(i: StatusInput): string {
  const t = T[i.lang];
  const ep = i.health.endpoint;
  const s = i.settings;
  const row = (k: string, v: string) => `<tr><th>${esc(k)}</th><td>${v}</td></tr>`;
  const badge = (ok: boolean, text: string) => `<span class="badge ${ok ? "ok" : "bad"}">${esc(text)}</span>`;
  // API-nøgler vises aldrig; den rå status er den samme som /api/health.
  const raw = JSON.stringify(i.health, null, 2);

  return `<!doctype html>
<html lang="${i.lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="refresh" content="30">
<title>${esc(t.title)}</title>
<style>
  :root { color-scheme: light dark; --bg:#f7f7f8; --card:#fff; --fg:#1b1b1f; --muted:#5f6368; --line:#e3e3e8; --ok:#0f7b3f; --okbg:#e6f4ea; --bad:#b3261e; --badbg:#fce8e6; }
  @media (prefers-color-scheme: dark) { :root { --bg:#121214; --card:#1d1d21; --fg:#ececf1; --muted:#a0a0aa; --line:#33333a; --ok:#7fd8a2; --okbg:#1c3427; --bad:#f2b8b5; --badbg:#3c1f1d; } }
  body { margin:0; background:var(--bg); color:var(--fg); font:15px/1.5 -apple-system, "Segoe UI", system-ui, sans-serif; }
  main { max-width:720px; margin:0 auto; padding:32px 16px; }
  h1 { font-size:22px; margin:0 0 4px; } h2 { font-size:15px; margin:24px 0 8px; color:var(--muted); text-transform:uppercase; letter-spacing:.04em; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px; }
  .lead { color:var(--muted); margin:0 0 20px; }
  table { width:100%; border-collapse:collapse; } th, td { text-align:left; padding:6px 8px; border-bottom:1px solid var(--line); vertical-align:top; } th { width:38%; color:var(--muted); font-weight:500; }
  tr:last-child th, tr:last-child td { border-bottom:0; }
  code, pre { font-family: ui-monospace, SFMono-Regular, Consolas, monospace; font-size:13px; }
  pre { margin:8px 0 0; padding:12px; background:var(--bg); border-radius:8px; overflow-x:auto; }
  .badge { display:inline-block; padding:1px 8px; border-radius:999px; font-size:13px; font-weight:600; }
  .ok { color:var(--ok); background:var(--okbg); } .bad { color:var(--bad); background:var(--badbg); }
  .note { border-left:3px solid var(--line); padding-left:12px; color:var(--muted); }
</style>
</head>
<body>
<main>
  <h1>${esc(t.title)} ${badge(true, t.running)}</h1>
  <p class="lead">${esc(t.officeOnly)}</p>

  <div class="card"><table>
    ${row(t.version, esc(i.version))}
    ${row(t.endpoint, `${esc(s.endpoints.find((e) => e.id === ep.id)?.name ?? ep.id)} – <code>${esc(ep.baseUrl)}</code> ${ep.reachable ? badge(true, t.reachable) : badge(false, t.unreachable)}${!ep.reachable && ep.error ? `<br><span class="note">${esc(ep.error)}</span>` : ""}`)}
    ${row(t.endpoints, s.endpoints.map((e) => `${esc(e.name)}: <code>${esc(e.baseUrl)}</code>`).join("<br>"))}
    ${row(t.addin, i.addinBuilt ? `${esc(t.addinOk)} <code>/taskpane.html</code>` : badge(false, t.addinMissing))}
  </table></div>

  <h2>${esc(t.policy)}</h2>
  <div class="card"><table>
    ${row(t.policy, esc(t.policySource[i.policySource] ?? i.policySource))}
    ${row(t.allowedHosts, i.policy.allowedHosts.map((h) => `<code>${esc(h)}</code>`).join(", "))}
    ${row(t.webSearch, !i.policy.webSearch ? esc(t.forbidden) : esc(s.webSearch ? t.on : t.off))}
    ${row(t.auditLog, esc(s.auditLog ? t.on : t.off))}
    ${row(t.config, `<code>${esc(i.configFile)}</code>`)}
  </table></div>

  <h2>${esc(t.api)}</h2>
  <div class="card">
    <p class="note">${esc(t.apiNote)}</p>
    <pre>${esc(raw)}</pre>
  </div>
</main>
</body>
</html>`;
}
