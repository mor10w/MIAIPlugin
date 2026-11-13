import type { Policy } from "@miai/shared";

/**
 * Content-Security-Policy for task pane (#28, opfølgning på ADR-0017).
 *
 * Politikken afgør allerede i koden, hvilke hosts MIAI må kalde. CSP'en lader browseren håndhæve
 * det samme, så heller ikke kode, der er kommet ind ad en anden vej (fx en fejl i visningen af
 * modeloutput), kan sende dokumentindhold ud: netværkskald (`connect-src`), billeder (`img-src`),
 * medier, plugins og formularer er begrænset til add-in'ets egen adresse og politikkens hosts.
 *
 * Bevidst ikke `script-src`/`default-src`: Office.js hentes fra Microsofts CDN og indlæser selv
 * mere kode, og kodetilstandens sandkasse bruger `new Function`. En for stram script-politik kan
 * ikke afprøves i rigtig Office herfra, og det er dataudslip – ikke scriptkilder – der er målet.
 */

/** CSP-kilder for politikkens hosts: http og https på alle porte ("*.domæne" og IPv6 understøttes). */
export function cspHostSources(allowedHosts: string[]): string[] {
  const out = new Set<string>();
  for (const raw of allowedHosts) {
    const h = raw
      .trim()
      .toLowerCase()
      .replace(/^\[|\]$/g, "");
    if (h === "*") return ["http:", "https:"];
    const host = h.includes(":") ? (/^[0-9a-f:.]+$/.test(h) ? `[${h}]` : "") : /^(\*\.)?[a-z0-9]([a-z0-9.-]*[a-z0-9])?$/.test(h) ? h : "";
    if (!host) continue; // ugyldigt navn: kan ikke udtrykkes i en CSP og tillades derfor ikke
    out.add(`http://${host}:*`);
    out.add(`https://${host}:*`);
  }
  return [...out];
}

/**
 * CSP-teksten. Med bridge går alle kald til bridgen selv (samme adresse), så ingen andre hosts
 * er nødvendige; i direkte tilstand tilføjes politikkens hosts.
 */
export function buildCsp(mode: "bridge" | "direct", policy: Pick<Policy, "allowedHosts">): string {
  const connect = ["'self'", ...(mode === "direct" ? cspHostSources(policy.allowedHosts) : [])].join(" ");
  return [
    `connect-src ${connect}`,
    "img-src 'self' data: blob:",
    "media-src 'self' data: blob:",
    "object-src 'none'",
    "form-action 'none'",
    "base-uri 'self'",
  ].join("; ");
}

const connectSources = (csp: string) => new Set((/connect-src ([^;]*)/.exec(csp)?.[1] ?? "").split(/\s+/).filter(Boolean));

let applied: string | undefined;

/** Den CSP, task pane kører med lige nu (til test og fejlsøgning). */
export const appliedCsp = (): string | undefined => applied;

/**
 * Sætter CSP'en som <meta> i dokumentet. En CSP kan kun strammes, mens siden er åben: returnerer
 * "reload", hvis den nye politik tillader hosts, den gældende CSP spærrer – så skal task pane
 * genindlæses, før de kan kaldes.
 */
export function applyCsp(csp: string, doc: Document = document): "ok" | "reload" {
  if (csp === applied) return "ok";
  if (applied) {
    const now = connectSources(applied);
    if ([...connectSources(csp)].some((src) => !now.has(src) && !now.has("http:"))) return "reload";
  }
  const meta = doc.createElement("meta");
  meta.httpEquiv = "Content-Security-Policy";
  meta.content = csp;
  doc.head.appendChild(meta);
  applied = csp;
  return "ok";
}

const RELOAD_FLAG = "miai.policyReload";

/**
 * Efter at politikken er gemt: stram CSP'en med det samme, eller genindlæs task pane, hvis der er
 * kommet hosts til (en CSP kan ikke udvides, mens siden er åben). Returnerer true ved genindlæsning.
 */
export function syncCspAfterSave(mode: "bridge" | "direct", policy: Pick<Policy, "allowedHosts">): boolean {
  if (applyCsp(buildCsp(mode, policy)) === "ok") return false;
  try {
    sessionStorage.setItem(RELOAD_FLAG, "1");
  } catch {
    // Uden sessionStorage åbner task pane bare på første fane efter genindlæsningen.
  }
  window.location.reload();
  return true;
}

/** Er task pane lige blevet genindlæst, fordi politikken blev ændret? `clear` nulstiller mærket. */
export function policyReloadPending(clear = false): boolean {
  try {
    const pending = sessionStorage.getItem(RELOAD_FLAG) === "1";
    if (pending && clear) sessionStorage.removeItem(RELOAD_FLAG);
    return pending;
  } catch {
    return false;
  }
}
