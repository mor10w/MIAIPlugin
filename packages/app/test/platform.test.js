import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import * as fs from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import selfsigned from "selfsigned";
import { macPlatform, manifestId, thumbprint, WIN_WEF_KEY, winPlatform } from "../platform.js";

/** MIAI App pr. platform (#5): Office-installation og betroelse af certifikatet. */

const MANIFEST = `<OfficeApp><Id>fdac46a3-0000-4e84-b704-78a669531789</Id><SourceLocation DefaultValue="https://localhost:45000/taskpane.html?token=x"/></OfficeApp>`;
const tmp = () => mkdtempSync(join(tmpdir(), "miai-app-"));

/** Falsk spawnSync, der husker kaldene og svarer med det, testen bestemmer. */
function fakeRun(reply = () => ({ status: 0, stdout: "", stderr: "" })) {
  const calls = [];
  const run = (cmd, args) => {
    calls.push([cmd, ...args]);
    return reply(cmd, args);
  };
  return { run, calls };
}

describe("macOS", () => {
  it("lægger manifestet i Word, Excel og PowerPoint og opdager ny port/token som 'stale'", () => {
    const home = tmp();
    const dir = tmp();
    const file = join(dir, "manifest.xml");
    writeFileSync(file, MANIFEST);
    const p = macPlatform({ home, fs });
    expect(p.officeInstallState(file)).toBe("missing");
    p.installInOffice(file);
    for (const a of ["Word", "Excel", "Powerpoint"])
      expect(
        readFileSync(join(home, "Library", "Containers", `com.microsoft.${a}`, "Data", "Documents", "wef", `miai-${manifestId(MANIFEST)}.xml`), "utf8"),
      ).toBe(MANIFEST);
    expect(p.officeInstallState(file)).toBe("ok");
    writeFileSync(file, MANIFEST.replace("45000", "45001"));
    expect(p.officeInstallState(file)).toBe("stale");
  });

  it("betror certifikatet i brugerens login-nøglering (uden admin)", () => {
    const { run, calls } = fakeRun();
    const p = macPlatform({ run, home: "/Users/a" });
    expect(p.trustCert("/c.crt")).toEqual({ ok: true, message: "" });
    expect(calls[0]).toEqual([
      "security",
      "add-trusted-cert",
      "-d",
      "-r",
      "trustRoot",
      "-k",
      join("/Users/a", "Library", "Keychains", "login.keychain-db"),
      "/c.crt",
    ]);
    expect(p.trayIcon).toEqual({ file: "tray.png", template: true });
  });
});

describe("Windows", () => {
  it("registrerer manifestets sti i HKCU (ingen admin) og læser tilstanden derfra", () => {
    const dir = tmp();
    const file = join(dir, "manifest.xml");
    writeFileSync(file, MANIFEST);
    const id = manifestId(MANIFEST);
    let registered = "";
    const { run, calls } = fakeRun((cmd, args) => {
      if (args[0] === "add") registered = args[args.indexOf("/d") + 1];
      if (args[0] === "query")
        return registered ? { status: 0, stdout: `\n${WIN_WEF_KEY}\n    ${id}    REG_SZ    ${registered}\n` } : { status: 1, stdout: "" };
      return { status: 0, stdout: "" };
    });
    const p = winPlatform({ run, fs });
    expect(p.officeInstallState(file)).toBe("missing");
    p.installInOffice(file);
    expect(calls.find((c) => c[1] === "add")).toEqual(["reg", "add", WIN_WEF_KEY, "/v", id, "/t", "REG_SZ", "/d", file, "/f"]);
    expect(p.officeInstallState(file)).toBe("ok");
    // En anden sti (fx gammel installation) er "stale".
    registered = "C:\\gammel\\manifest.xml";
    expect(p.officeInstallState(file)).toBe("stale");
  });

  it("fejler med en forklaring, hvis registreringsdatabasen afviser", () => {
    const dir = tmp();
    const file = join(dir, "manifest.xml");
    writeFileSync(file, MANIFEST);
    const p = winPlatform({ run: () => ({ status: 1, stderr: "Adgang nægtet" }), fs });
    expect(() => p.installInOffice(file)).toThrow(/Adgang nægtet/);
  });

  it("betror certifikatet i brugerens rodlager og genkender det på tommelfingeraftrykket", async () => {
    const dir = tmp();
    mkdirSync(dir, { recursive: true });
    const pems = await selfsigned.generate([{ name: "commonName", value: "localhost" }]);
    const cert = join(dir, "localhost.crt");
    writeFileSync(cert, pems.cert);
    const tp = thumbprint(pems.cert);
    expect(tp).toMatch(/^[0-9A-F]{40}$/);
    const { run, calls } = fakeRun((cmd, args) => ({ status: args[1] === "-verifystore" && args[3] === tp ? 0 : 0, stdout: "" }));
    const p = winPlatform({ run, fs });
    expect(p.trustCert(cert).ok).toBe(true);
    expect(calls[0]).toEqual(["certutil", "-user", "-addstore", "Root", cert]);
    expect(p.isCertTrusted(cert)).toBe(true);
    expect(calls[1]).toEqual(["certutil", "-user", "-verifystore", "Root", tp]);
    expect(p.trayIcon).toEqual({ file: "tray-color.png", template: false });
    expect(p.fileManager).toBe("Stifinder");
  });
});
