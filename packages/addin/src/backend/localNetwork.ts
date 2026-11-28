import type { LocalNetworkAccess } from "@miai/shared";

/**
 * Lokal netværksadgang (Local Network Access, #74). Chrome 142+ og Edge 143+ kræver brugerens
 * tilladelse, før en side fra internettet må kalde maskinen selv (loopback) eller det lokale net.
 * Det rammer direkte tilstand i Office på nettet: add-in'et er hostet offentligt og kalder fx
 * http://127.0.0.1:11434. Er tilladelsen afvist, fejler fetch med en netværksfejl (TypeError), som
 * ikke kan skelnes fra en lukket port – derfor spørges Permissions API'et, når et kald er fejlet.
 *
 * Bekræftet i kilderne (WICG-specifikationen og Chromiums IDL), men ikke afprøvet i en rigtig browser:
 *  - tilladelserne hedder "loopback-network" og "local-network";
 *  - en afvist tilladelse giver en netværksfejl, også for kald med mode "no-cors".
 * Ubekræftet, og derfor kodet defensivt:
 *  - De første versioner med kravet (efter sigende Chrome 142–144) kender kun det samlede navn
 *    "local-network-access". Det spørges der ikke om: Chromes vejledning advarer efter sigende om,
 *    at det kan få ældre versioner til at gå ned. De versioner opfører sig derfor som browsere
 *    uden API'et.
 *  - document.featurePolicy (kun Chromium, ikke en standard) bruges til at se, om siden, add-in'et
 *    vises i, giver tilladelsen videre. Kender browseren ikke funktionen, regnes den som givet videre.
 * Tilstanden siger ikke, om browseren håndhæver kravet for netop denne side; "denied" er dog altid
 * et udtrykkeligt nej (fra brugeren, en politik eller siden uden om add-in'et).
 */

export type AddressSpace = "loopback" | "local";

/**
 * Adresserummet for en URL, når det fremgår af selve adressen (localhost, IP-adresse, .local).
 * Andre navne giver undefined: de kan pege på internettet, og så gælder kravet ikke.
 */
export function addressSpace(url: string): AddressSpace | undefined {
  let host: string;
  try {
    host = new URL(url).hostname.toLowerCase();
  } catch {
    return undefined;
  }
  if (host === "localhost" || host.endsWith(".localhost") || host === "[::1]") return "loopback";
  const v4 = /^(\d+)\.(\d+)\.\d+\.\d+$/.exec(host);
  if (v4) {
    const a = Number(v4[1]);
    const b = Number(v4[2]);
    if (a === 127) return "loopback";
    const local = a === 10 || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || (a === 169 && b === 254) || (a === 100 && b >= 64 && b <= 127);
    return local ? "local" : undefined;
  }
  // IPv6: fc00::/7 (unikke lokale) og fe80::/10 (link-lokale).
  if (/^\[(f[cd]|fe[89ab])/.test(host) || host.endsWith(".local")) return "local";
  return undefined;
}

/** Jo lavere tal, jo mere lokalt. Kravet gælder kald til et mere lokalt adresserum end sidens eget. */
const rank = (url: string): number => ({ loopback: 0, local: 1, public: 2 })[addressSpace(url) ?? "public"];

type FeaturePolicy = { features?: () => string[]; allowsFeature?: (name: string) => boolean };

/** Tilladelsens tilstand – eller undefined, når browseren ikke kender den (ældre browsere, Office's webviews, Firefox, Safari). */
async function permissionState(name: string): Promise<string | undefined> {
  try {
    if (typeof navigator === "undefined" || typeof navigator.permissions?.query !== "function") return undefined;
    // Et ukendt navn afvises med en TypeError.
    return (await navigator.permissions.query({ name: name as PermissionName })).state;
  } catch {
    return undefined;
  }
}

/** Vises add-in'et i en ramme, som ikke har fået tilladelsen givet videre? Kan det ikke afgøres: nej. */
function withheldByFrame(name: string): boolean {
  try {
    if (window.top === window.self) return false;
    const policy = (document as { featurePolicy?: FeaturePolicy }).featurePolicy;
    if (!policy?.allowsFeature || !policy.features?.().includes(name)) return false;
    return !policy.allowsFeature(name);
  } catch {
    return false;
  }
}

/**
 * Står browserens tilladelse i vejen for et kald til `url`? Kaldes, når et kald er fejlet.
 * undefined = nej eller ukendt (tilladt, ikke en lokal adresse, eller browseren har ikke API'et),
 * og så behandles fejlen som hidtil. "prompt" gives kun, når kravet gælder for siden: en side på
 * maskinen selv (bridge, udvikling på https://localhost) må kalde lokale adresser uden tilladelse.
 */
export async function localNetworkAccess(url: string, page: string = window.location.href): Promise<LocalNetworkAccess | undefined> {
  const space = addressSpace(url);
  if (!space) return undefined;
  const name = space === "loopback" ? "loopback-network" : "local-network";
  const state = await permissionState(name);
  if (state === "denied") return withheldByFrame(name) ? "frame" : "denied";
  if (state === "prompt" && rank(url) < rank(page)) return "prompt";
  return undefined;
}
