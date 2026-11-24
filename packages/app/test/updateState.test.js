import { describe, expect, it } from "vitest";
import { macDownloadUrl, UPDATE_FEED, updateMenuItems, updateMode, updateNotification } from "../updateState.js";

describe("opdatering fra menuen", () => {
  it("updateMode: Windows opdaterer selv, Mac giver besked (indtil #67), udviklerudgaven følger sin kanal", () => {
    expect(updateMode("win32", "0.1.20")).toBe("auto");
    expect(updateMode("win32", "0.1.20-dev.9")).toBe("auto-dev");
    expect(updateMode("darwin", "0.1.20")).toBe("notify");
    expect(updateMode("darwin", "0.1.20-dev.9")).toBe("browser-dev");
    expect(updateMode("linux", "0.1.20")).toBe("none");
  });

  it("macDownloadUrl: dmg'en for udgivelsen og arkitekturen på feedet, ellers null", () => {
    expect(macDownloadUrl("0.1.21", "arm64")).toBe(`${UPDATE_FEED}MIAI-0.1.21-arm64.dmg`);
    expect(macDownloadUrl("0.1.21", "x64")).toBe(`${UPDATE_FEED}MIAI-0.1.21-x64.dmg`);
    for (const v of ["0.1.21-dev.3", "../x", "0.1", "", undefined]) expect(macDownloadUrl(v, "arm64"), String(v)).toBeNull();
    expect(macDownloadUrl("0.1.21", "ia32")).toBeNull();
    expect(macDownloadUrl("0.1.21", "arm64", "http://usikker.example/")).toBeNull();
  });

  it("menuen viser altid et opdateringspunkt, der passer til tilstanden", () => {
    expect(updateMenuItems({ status: "idle" })).toEqual([{ label: "Søg efter opdateringer…", action: "check" }]);
    expect(updateMenuItems({ status: "latest" })).toEqual([{ label: "Søg efter opdateringer…", action: "check" }]);
    expect(updateMenuItems({ status: "error", error: "x" })).toEqual([{ label: "Søg efter opdateringer…", action: "check" }]);
    expect(updateMenuItems({ status: "checking" })).toEqual([{ label: "Søger efter opdateringer…", enabled: false }]);
    expect(updateMenuItems({ status: "downloading", version: "0.1.21" })).toEqual([{ label: "Henter MIAI 0.1.21…", enabled: false }]);
    expect(updateMenuItems({ status: "ready", version: "0.1.21" })).toEqual([{ label: "Genstart og opdater til 0.1.21", action: "install" }]);
    expect(updateMenuItems({ status: "available", version: "0.1.21", url: "u" })).toEqual([{ label: "Hent MIAI 0.1.21…", action: "open", url: "u" }]);
    expect(updateMenuItems(undefined)[0].action).toBe("check");
  });

  it("besked om en ny udgave én gang pr. version, også fra baggrundstjekket", () => {
    const notified = new Set();
    const ready = updateNotification({ status: "ready", version: "0.1.21" }, { notified });
    expect(ready).toMatchObject({ key: "ready:0.1.21", action: "install", title: "MIAI er klar til at opdatere" });
    notified.add(ready.key);
    expect(updateNotification({ status: "ready", version: "0.1.21" }, { notified })).toBeNull();
    expect(updateNotification({ status: "ready", version: "0.1.22" }, { notified })?.key).toBe("ready:0.1.22");

    const avail = updateNotification({ status: "available", version: "0.1.21", url: "u" }, { notified });
    expect(avail).toMatchObject({ key: "available:0.1.21", action: "open", url: "u" });
  });

  it("'nyeste version', fejl og 'henter' siges kun til, når brugeren selv søgte", () => {
    for (const state of [{ status: "latest" }, { status: "error", error: "intet net" }, { status: "downloading", version: "0.1.21" }, { status: "checking" }])
      expect(updateNotification(state, { manual: false }), state.status).toBeNull();
    expect(updateNotification({ status: "latest" }, { manual: true, current: "0.1.20" })?.body).toBe("Du har den nyeste version (0.1.20).");
    expect(updateNotification({ status: "error", error: "intet net" }, { manual: true })?.body).toBe("intet net");
    expect(updateNotification({ status: "downloading", version: "0.1.21" }, { manual: true })?.body).toContain("0.1.21 hentes i baggrunden");
    expect(updateNotification({ status: "checking" }, { manual: true })).toBeNull();
  });
});
