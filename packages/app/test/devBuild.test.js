import { describe, expect, it } from "vitest";
import { DEV_FEED, devBuildNumber, devBuildSupported, devMenuState, fetchDevInfo, parseDevInfo, platformKey } from "../devBuild.js";

const INFO = { version: "0.1.0-dev.42", file: "MIAI-0.1.0-dev.42-arm64.dmg", commit: "66e2542abcdef0123456", date: "2026-11-13T19:43:00Z" };
const WIN = "win32-x64";
/** dev.json, som dev-build.yml skriver det efter #40: ét filnavn pr. platform, og `file` til ældre Mac-udgaver. */
const BOTH = { ...INFO, files: { "darwin-arm64": "MIAI-0.1.0-dev.42-arm64.dmg", "win32-x64": "MIAI-Setup-0.1.0-dev.42-x64.exe" } };

describe("udviklerudgaven (ADR-0016, tillæg)", () => {
  it("byggenummeret læses kun fra udviklerversioner", () => {
    expect(devBuildNumber("0.1.0-dev.42")).toBe(42);
    expect(devBuildNumber("0.1.0")).toBeNull();
    expect(devBuildNumber("0.1.0-beta.1")).toBeNull();
    expect(devBuildNumber(undefined)).toBeNull();
  });

  it("tilbydes kun, hvor den bygges: Mac på Apple Silicon og Windows på x64 (#40)", () => {
    expect(devBuildSupported("darwin", "arm64")).toBe(true);
    expect(devBuildSupported("win32", "x64")).toBe(true);
    expect(devBuildSupported("darwin", "x64")).toBe(false);
    expect(devBuildSupported("win32", "arm64")).toBe(false);
    expect(devBuildSupported("linux", "x64")).toBe(false);
    expect(platformKey("win32", "x64")).toBe(WIN);
  });

  it("dev.json giver en adresse i feedets egen mappe", () => {
    expect(parseDevInfo(INFO)).toEqual({ version: "0.1.0-dev.42", url: `${DEV_FEED}MIAI-0.1.0-dev.42-arm64.dmg`, commit: "66e2542abcde", date: "2026-11-13" });
  });

  it("dev.json med files: hver platform får sin egen fil (#40)", () => {
    expect(parseDevInfo(BOTH, DEV_FEED, "darwin-arm64")?.url).toBe(`${DEV_FEED}MIAI-0.1.0-dev.42-arm64.dmg`);
    expect(parseDevInfo(BOTH, DEV_FEED, WIN)?.url).toBe(`${DEV_FEED}MIAI-Setup-0.1.0-dev.42-x64.exe`);
    // Uden et Windows-byg i feedet (fx et ældre feed med kun `file`) tilbydes Windows intet.
    expect(parseDevInfo(INFO, DEV_FEED, WIN)).toBeNull();
    expect(parseDevInfo({ ...BOTH, files: { "darwin-arm64": BOTH.files["darwin-arm64"] } }, DEV_FEED, WIN)).toBeNull();
    // Med files bruges `file` ikke som reserve, og en ukendt platform får intet.
    expect(parseDevInfo({ ...BOTH, files: {} }, DEV_FEED, "darwin-arm64")).toBeNull();
    expect(parseDevInfo(BOTH, DEV_FEED, "linux-x64")).toBeNull();
    expect(parseDevInfo(BOTH, DEV_FEED, "__proto__")).toBeNull();
  });

  it("hver platform får kun sin egen filtype og et rent filnavn i feedets mappe (#40)", () => {
    const withWin = (file) => ({ ...BOTH, files: { ...BOTH.files, [WIN]: file } });
    for (const file of ["https://evil.example/x.exe", "../x.exe", "//evil.example/x.exe", "a/b.exe", "x.dmg", "x.exe.bat", "x.msi", "", 5])
      expect(parseDevInfo(withWin(file), DEV_FEED, WIN), String(file)).toBeNull();
    expect(parseDevInfo({ ...BOTH, files: { ...BOTH.files, "darwin-arm64": "x.exe" } }, DEV_FEED, "darwin-arm64")).toBeNull();
  });

  it("et feed kan ikke sende brugeren et andet sted hen eller til noget, der ikke er en .dmg", () => {
    for (const file of ["https://evil.example/x.dmg", "../x.dmg", "//evil.example/x.dmg", "x.exe", "a/b.dmg", "", 5])
      expect(parseDevInfo({ ...INFO, file }), String(file)).toBeNull();
    expect(parseDevInfo({ ...INFO, version: "9.9.9" })).toBeNull();
    expect(parseDevInfo(INFO, "http://usikker.example/dev/")).toBeNull();
    expect(parseDevInfo(null)).toBeNull();
    expect(parseDevInfo("tekst")).toBeNull();
  });

  it("menupunktet: hent en nyere, vis at den er installeret, eller intet uden feed", () => {
    const info = parseDevInfo(INFO);
    expect(devMenuState("0.1.0-dev.41", info)).toEqual({
      kind: "available",
      label: "Hent udviklerudgave 0.1.0-dev.42 (2026-11-13)…",
      url: `${DEV_FEED}MIAI-0.1.0-dev.42-arm64.dmg`,
    });
    // Fra en almindelig udgivelse kan man altid skifte til udviklerudgaven.
    expect(devMenuState("0.1.0", info).kind).toBe("available");
    expect(devMenuState("0.1.0-dev.42", info)).toEqual({ kind: "current", label: "Udviklerudgave 0.1.0-dev.42 er installeret" });
    // Feedet er bagud (fx mens en ny bygges): den installerede er nyere og tilbydes ikke nedgraderet.
    expect(devMenuState("0.1.0-dev.50", info)).toEqual({ kind: "current", label: "Udviklerudgave 0.1.0-dev.50 er installeret" });
    expect(devMenuState("0.1.0", null)).toEqual({ kind: "none" });
  });

  it("fetchDevInfo: henter dev.json og giver null ved fejl", async () => {
    const ok = async (url) => ({ ok: true, json: async () => INFO, url: String(url) });
    expect((await fetchDevInfo(DEV_FEED, ok)).version).toBe("0.1.0-dev.42");
    const both = async () => ({ ok: true, json: async () => BOTH });
    expect((await fetchDevInfo(DEV_FEED, both, WIN)).url).toBe(`${DEV_FEED}MIAI-Setup-0.1.0-dev.42-x64.exe`);
    expect(await fetchDevInfo(DEV_FEED, async () => ({ ok: false }))).toBeNull();
    expect(
      await fetchDevInfo(DEV_FEED, async () => {
        throw new Error("intet net");
      }),
    ).toBeNull();
    expect(await fetchDevInfo(DEV_FEED, async () => ({ ok: true, json: async () => ({ version: "x" }) }))).toBeNull();
  });
});
