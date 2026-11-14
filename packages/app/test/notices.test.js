import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { collectPackages, describePackage, packageRootOf, renderNotices, runtimeClosure } from "../../../scripts/third-party-notices.mjs";

// Et lille, opdigtet node_modules: app → a → b (indlejret i a), og en dev-afhængighed, der ikke må komme med.
const root = mkdtempSync(join(tmpdir(), "miai-notices-"));
afterAll(() => rmSync(root, { recursive: true, force: true }));

function pkg(dir, json, files = {}) {
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, "package.json"), JSON.stringify(json));
  for (const [name, text] of Object.entries(files)) writeFileSync(join(dir, name), text);
}

const app = join(root, "app");
const nm = join(root, "node_modules");
pkg(app, {
  name: "@miai/app",
  version: "1.0.0",
  dependencies: { a: "1", "@miai/shared": "1" },
  devDependencies: { værktøj: "1" },
  optionalDependencies: { "kun-windows": "1" },
});
pkg(
  join(nm, "a"),
  {
    name: "a",
    version: "1.2.3",
    license: "MIT",
    author: "Ada <ada@example.com> (https://example.com)",
    repository: { url: "git+https://example.com/a.git" },
    dependencies: { b: "2" },
  },
  { LICENSE: "MIT\r\nCopyright Ada\r\n", "NOTICE.txt": "Bemærk" },
);
pkg(join(nm, "a", "node_modules", "b"), { name: "b", version: "2.0.0", licenses: [{ type: "ISC" }] });
mkdirSync(join(nm, "a", "dist", "esm"), { recursive: true });
writeFileSync(join(nm, "a", "dist", "esm", "package.json"), '{"type":"module"}');
pkg(join(nm, "@miai", "shared"), { name: "@miai/shared", version: "1.0.0", license: "MIT" }, { LICENSE: "MIT" });
pkg(join(nm, "værktøj"), { name: "værktøj", version: "1.0.0", license: "MIT" }, { LICENSE: "MIT" });

describe("tredjepartslicenser", () => {
  it("finder pakkens rod ud fra en fil, også forbi en indre package.json uden navn", () => {
    expect(packageRootOf(join(nm, "a", "dist", "esm", "index.js?v=1"))).toBe(join(nm, "a"));
    expect(packageRootOf(join(nm, "a", "node_modules", "b", "index.js"))).toBe(join(nm, "a", "node_modules", "b"));
    expect(packageRootOf(join(app, "main.js"))).toBeNull();
    expect(packageRootOf("\0virtual:noget")).toBeNull();
  });

  it("læser licens, kilde og tekst (LICENSE og NOTICE) uden e-mail i forfatterfeltet", () => {
    expect(describePackage(join(nm, "a"))).toEqual({
      name: "a",
      version: "1.2.3",
      license: "MIT",
      author: "Ada",
      repository: "https://example.com/a",
      text: "MIT\nCopyright Ada\n\nBemærk",
    });
    expect(describePackage(join(nm, "a", "node_modules", "b"))).toMatchObject({ license: "ISC", text: "" });
    expect(describePackage(join(nm, "a", "dist", "esm"))).toBeNull();
  });

  it("følger kun det, der udleveres: dependencies transitivt, ikke devDependencies og ikke manglende valgfrie pakker", () => {
    const names = collectPackages(runtimeClosure(app)).map((p) => `${p.name}@${p.version}`);
    // MIAI's egne pakker er ikke tredjepart.
    expect(names).toEqual(["a@1.2.3", "b@2.0.0"]);
  });

  it("filen nævner hver pakke med licens og siger det ligeud, når en pakke ingen licensfil har", () => {
    const text = renderNotices(collectPackages(runtimeClosure(app)), "Indledning.\n");
    expect(text).toContain("Antal pakker: 2");
    expect(text).toContain("a 1.2.3\nLicens: MIT\nKilde: https://example.com/a");
    expect(text).toContain("Copyright Ada");
    expect(text).toContain("b 2.0.0\nLicens: ISC");
    expect(text).toContain("Pakken indeholder ingen licensfil. Licens ifølge pakken: ISC.");
    expect(text).not.toContain("ada@example.com");
  });
});
