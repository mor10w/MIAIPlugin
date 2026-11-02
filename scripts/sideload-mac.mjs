// Sideloader MIAI-manifestet i Word, Excel og PowerPoint på macOS.
// Kør: npm run sideload:mac   (genstart derefter Office-appen)
import { copyFileSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const manifest = new URL("../packages/addin/manifest.xml", import.meta.url);
for (const app of ["Word", "Excel", "Powerpoint"]) {
  const dir = join(homedir(), "Library", "Containers", `com.microsoft.${app}`, "Data", "Documents", "wef");
  mkdirSync(dir, { recursive: true });
  copyFileSync(manifest, join(dir, "miai-manifest.xml"));
  console.log(`✓ ${app}: ${dir}`);
}
