import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const read = (f) => readFileSync(new URL(`../${f}`, import.meta.url), "utf8");

/** Lokale moduler, som en fil importerer ("./x.js"), rekursivt fra main.js. */
function localImports(file, seen = new Set()) {
  for (const m of read(file).matchAll(/(?:from|import)\s*\(?\s*"\.\/([^"]+)"/g)) {
    if (seen.has(m[1])) continue;
    seen.add(m[1]);
    localImports(m[1], seen);
  }
  return seen;
}

describe("pakning", () => {
  it("alle lokale moduler, appen importerer, kommer med i den installerede app (build.files)", () => {
    const files = JSON.parse(read("package.json")).build.files;
    const missing = [...localImports("main.js")].filter((f) => !files.includes(f));
    expect(missing).toEqual([]);
    expect(localImports("main.js")).toContain("updateState.js");
  });
});
