// Sideloader et MIAI-manifest i Word, Excel og PowerPoint på macOS.
//   npm run sideload:mac                                   → manifest.dev.xml (direkte, localhost:3000)
//   npm run sideload:mac -- packages/addin/manifest.xml    → bridge-varianten
//   npm run sideload:mac -- ~/Downloads/manifest.xml       → hostet variant (fx fra GitHub Pages)
// Genstart derefter Office-appen.
import { copyFileSync, mkdirSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

const manifest = resolve(process.argv[2] ?? new URL("../packages/addin/manifest.dev.xml", import.meta.url).pathname);
const id = readFileSync(manifest, "utf8").match(/<Id>([^<]+)<\/Id>/)?.[1];
if (!id) throw new Error(`Fandt ikke <Id> i ${manifest}`);
for (const app of ["Word", "Excel", "Powerpoint"]) {
  const dir = join(homedir(), "Library", "Containers", `com.microsoft.${app}`, "Data", "Documents", "wef");
  mkdirSync(dir, { recursive: true });
  copyFileSync(manifest, join(dir, `miai-${id}.xml`));
  console.log(`✓ ${app}: ${dir}`);
}
