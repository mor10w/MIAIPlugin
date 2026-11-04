// Det, der er forskelligt på macOS og Windows (#5): hvor Office finder add-in'et, hvordan
// certifikatet betros, og hvad Filhåndteringen hedder. main.js kender kun dette interface.
//
// Funktionerne tager `run` (spawnSync) og `fs` som parametre, så de kan testes uden at
// røre maskinens nøglering, registreringsdatabase eller Office.
import { spawnSync } from "node:child_process";
import * as nodeFs from "node:fs";
import { X509Certificate } from "node:crypto";
import { homedir } from "node:os";
import { join } from "node:path";

/** Word, Excel og PowerPoint (Outlook installeres fra fil – se BRUG.md). */
const MAC_OFFICE_APPS = ["Word", "Excel", "Powerpoint"];
/** Registreringsnøglen, Office på Windows læser sideloadede manifester fra (som office-addin-dev-settings). */
export const WIN_WEF_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";

/** Manifestets <Id>. */
export function manifestId(xml) {
  const id = xml.match(/<Id>([^<]+)<\/Id>/)?.[1];
  if (!id) throw new Error("Fandt ikke <Id> i manifestet");
  return id;
}

/** SHA-1-tommelfingeraftryk uden kolon (sådan som certutil og Windows' certifikatlager skriver det). */
export function thumbprint(certPem) {
  return new X509Certificate(certPem).fingerprint.replace(/:/g, "").toUpperCase();
}

/** macOS: manifestet lægges i hver Office-apps wef-mappe; certifikatet betros i login-nøgleringen. */
export function macPlatform({ run = spawnSync, fs = nodeFs, home = homedir() } = {}) {
  const wef = (a) => join(home, "Library", "Containers", `com.microsoft.${a}`, "Data", "Documents", "wef");
  return {
    name: "mac",
    fileManager: "Finder",
    /** Menulinje-ikonet er et skabelonbillede (sort glyph, som macOS farver efter temaet). */
    trayIcon: { file: "tray.png", template: true },
    installInOffice(manifestFile) {
      const id = manifestId(fs.readFileSync(manifestFile, "utf8"));
      for (const a of MAC_OFFICE_APPS) {
        fs.mkdirSync(wef(a), { recursive: true });
        fs.copyFileSync(manifestFile, join(wef(a), `miai-${id}.xml`));
      }
    },
    /** ok · stale (installeret med anden port/token) · missing */
    officeInstallState(manifestFile) {
      try {
        const xml = fs.readFileSync(manifestFile, "utf8");
        const files = MAC_OFFICE_APPS.map((a) => join(wef(a), `miai-${manifestId(xml)}.xml`));
        if (!files.every((f) => fs.existsSync(f))) return "missing";
        return files.every((f) => fs.readFileSync(f, "utf8") === xml) ? "ok" : "stale";
      } catch {
        return "missing";
      }
    },
    /** Egen login-nøglering: kræver ikke admin, og der vises ingen dialog. */
    trustCert(certFile) {
      const keychain = join(home, "Library", "Keychains", "login.keychain-db");
      const r = run("security", ["add-trusted-cert", "-d", "-r", "trustRoot", "-k", keychain, certFile], { encoding: "utf8" });
      return { ok: r.status === 0, message: (r.stderr ?? "").trim() };
    },
    isCertTrusted(certFile) {
      return run("security", ["verify-cert", "-c", certFile, "-p", "ssl", "-s", "localhost", "-q"], { encoding: "utf8" }).status === 0;
    },
  };
}

/**
 * Windows: Office læser stien til manifestet fra registreringsdatabasen (HKCU, kræver ikke admin);
 * filen opdateres af bridgen på stedet, så en ny port eller et nyt token gælder ved næste start af Office.
 * Certifikatet betros i brugerens eget rodlager – Windows viser en sikkerhedsdialog, brugeren godkender.
 */
export function winPlatform({ run = spawnSync, fs = nodeFs } = {}) {
  const query = (id) => run("reg", ["query", WIN_WEF_KEY, "/v", id], { encoding: "utf8", windowsHide: true });
  return {
    name: "win",
    fileManager: "Stifinder",
    /** Proceslinjen er lys eller mørk – brug app-ikonet i farver. */
    trayIcon: { file: "tray-color.png", template: false },
    installInOffice(manifestFile) {
      const id = manifestId(fs.readFileSync(manifestFile, "utf8"));
      const r = run("reg", ["add", WIN_WEF_KEY, "/v", id, "/t", "REG_SZ", "/d", manifestFile, "/f"], { encoding: "utf8", windowsHide: true });
      if (r.status !== 0) throw new Error(`Registreringsdatabasen afviste: ${(r.stderr ?? "").trim() || r.status}`);
    },
    officeInstallState(manifestFile) {
      try {
        const id = manifestId(fs.readFileSync(manifestFile, "utf8"));
        const r = query(id);
        if (r.status !== 0) return "missing";
        // Office peger på filen, som bridgen holder opdateret – kun en anden sti er "stale".
        return (r.stdout ?? "").toLowerCase().includes(manifestFile.toLowerCase()) ? "ok" : "stale";
      } catch {
        return "missing";
      }
    },
    trustCert(certFile) {
      const r = run("certutil", ["-user", "-addstore", "Root", certFile], { encoding: "utf8", windowsHide: true });
      return { ok: r.status === 0, message: (r.stdout || r.stderr || "").trim() };
    },
    isCertTrusted(certFile) {
      try {
        const tp = thumbprint(fs.readFileSync(certFile, "utf8"));
        return run("certutil", ["-user", "-verifystore", "Root", tp], { encoding: "utf8", windowsHide: true }).status === 0;
      } catch {
        return false;
      }
    },
  };
}

export function currentPlatform() {
  return process.platform === "win32" ? winPlatform() : macPlatform();
}
