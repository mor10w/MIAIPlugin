import { EventEmitter } from "node:events";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// updates.js med en falsk Electron og en falsk electron-updater: tilstande, menu og beskeder.
const env = { version: "0.1.20", notes: [], opened: [] };

vi.mock("electron", () => ({
  app: { isPackaged: true, getVersion: () => env.version },
  net: { fetch: async () => ({ ok: false }) },
  shell: { openExternal: async (u) => env.opened.push(u) },
  Notification: class extends EventEmitter {
    static isSupported() {
      return true;
    }
    constructor(o) {
      super();
      this.o = o;
    }
    show() {
      env.notes.push(this);
    }
  },
}));

class FakeUpdater extends EventEmitter {
  feed = null;
  installed = false;
  /** Versionen på feedet, eller null for ingen opdatering. */
  latest = null;
  setFeedURL(o) {
    this.feed = o;
  }
  quitAndInstall() {
    this.installed = true;
  }
  async checkForUpdates() {
    this.emit("checking-for-update");
    if (!this.latest) return void this.emit("update-not-available", {});
    this.emit("update-available", { version: this.latest });
    if (this.autoDownload) this.emit("update-downloaded", { version: this.latest });
  }
}
let updater;
vi.mock("electron-updater", () => ({
  default: {
    get autoUpdater() {
      return updater;
    },
  },
}));

const realPlatform = process.platform;
const realArch = process.arch;
const tick = () => new Promise((r) => setTimeout(r, 0));
const titles = () => env.notes.map((n) => n.o.title);

async function start(platform, version, arch = "x64") {
  Object.defineProperty(process, "platform", { value: platform });
  Object.defineProperty(process, "arch", { value: arch });
  env.version = version;
  const m = await import("../updates.js");
  await m.setupUpdates({ log: () => {}, onChange: () => {} });
  await tick();
  return m;
}

beforeEach(() => {
  vi.resetModules();
  updater = new FakeUpdater();
  env.notes = [];
  env.opened = [];
});
afterEach(() => {
  Object.defineProperty(process, "platform", { value: realPlatform });
  Object.defineProperty(process, "arch", { value: realArch });
});

describe("opdatering fra menuen (updates.js)", () => {
  it("Windows: søg, hent i baggrunden, besked når den er klar, og klik installerer", async () => {
    const m = await start("win32", "0.1.20");
    expect(updater.feed).toBeNull(); // almindelige udgaver bruger feedet i app-update.yml
    expect(updater.autoDownload).toBe(true);
    expect(env.notes).toHaveLength(0); // baggrundstjek uden opdatering: ingen besked
    expect(m.updateMenuItem().map((i) => i.label)).toEqual(["Søg efter opdateringer…"]);

    m.updateMenuItem()[0].click();
    await tick();
    expect(env.notes.at(-1).o.body).toBe("Du har den nyeste version (0.1.20).");

    updater.latest = "0.1.21";
    m.checkNow();
    await tick();
    expect(titles().slice(-2)).toEqual(["MIAI opdateres", "MIAI er klar til at opdatere"]);
    expect(m.updateMenuItem().map((i) => i.label)).toEqual(["Genstart og opdater til 0.1.21"]);

    // Baggrundstjek, mens den er klar: ingen ny besked, og menuen bliver ved "Genstart".
    const count = env.notes.length;
    await updater.checkForUpdates();
    expect(env.notes).toHaveLength(count);
    expect(m.updateMenuItem().map((i) => i.label)).toEqual(["Genstart og opdater til 0.1.21"]);

    env.notes.find((n) => n.o.title === "MIAI er klar til at opdatere").emit("click");
    expect(updater.installed).toBe(true);
  });

  it("Windows-udviklerudgaven følger kanalen dev (dev.yml under /dev/) uden at tillade nedgradering", async () => {
    await start("win32", "0.1.20-dev.5");
    expect(updater.feed).toEqual({ provider: "generic", url: "https://mor10w.github.io/MIAIPlugin/dev/", channel: "dev" });
    expect(updater.channel).toBeUndefined();
    expect(updater.allowDowngrade).toBeUndefined();
  });

  it("Mac: henter ikke selv (indtil #67), men giver besked én gang og åbner dmg'en for arkitekturen", async () => {
    updater.latest = "0.1.21";
    const m = await start("darwin", "0.1.20", "arm64");
    expect(updater.autoDownload).toBe(false);
    expect(titles()).toEqual(["Ny version af MIAI"]);
    expect(m.updateMenuItem().map((i) => i.label)).toEqual(["Hent MIAI 0.1.21…"]);

    await updater.checkForUpdates();
    expect(titles()).toEqual(["Ny version af MIAI"]); // ingen gentagen besked

    env.notes[0].emit("click");
    m.updateMenuItem()[0].click();
    expect(env.opened).toEqual([
      "https://mor10w.github.io/MIAIPlugin/updates/MIAI-0.1.21-arm64.dmg",
      "https://mor10w.github.io/MIAIPlugin/updates/MIAI-0.1.21-arm64.dmg",
    ]);
  });

  it("fejl siges kun til, når brugeren selv søgte", async () => {
    const m = await start("win32", "0.1.20");
    updater.checkForUpdates = async () => {
      throw new Error("intet net");
    };
    updater.emit("error", new Error("intet net"));
    expect(env.notes).toHaveLength(0);
    m.checkNow();
    await tick();
    expect(env.notes.at(-1).o).toEqual({ title: "Kunne ikke søge efter opdateringer", body: "intet net" });
  });
});
