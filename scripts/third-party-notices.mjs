// Samler licenserne for de tredjepartspakker, MIAI udleveres sammen med, i én tekstfil.
// MIT, BSD, Apache m.fl. kræver, at ophavsretsnotits og licenstekst følger med kopier af koden.
// Kildekoden henviser kun til pakkerne (package-lock.json); det er de byggede filer, der indeholder dem:
//   - add-in'et (packages/addin/dist): Vite-plugin'et i vite.config.ts kender præcis de moduler, der er pakket med.
//   - MIAI App: alt, appen afhænger af ved kørsel (dependencies, transitivt), fundet med lukningen nedenfor.
//
//   node scripts/third-party-notices.mjs app <ud-fil>   → licenser for MIAI Apps afhængigheder
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const LICENSE_FILE = /^(licen[sc]e|copying|notice|unlicense)([.-].*)?$/i;

/** Læser en pakkes package.json, eller null hvis mappen ikke er roden af en pakke. */
function readPackage(dir) {
  try {
    const pkg = JSON.parse(readFileSync(join(dir, "package.json"), "utf8"));
    // Undermapper kan have en package.json med kun {"type": "module"}; den er ikke pakkens rod.
    return typeof pkg.name === "string" && typeof pkg.version === "string" ? pkg : null;
  } catch {
    return null;
  }
}

/** Roden af den pakke i node_modules, som en fil hører til, eller null for filer uden for node_modules. */
export function packageRootOf(file) {
  const clean = file.replace(/^\0+/, "").split("?")[0];
  if (!/[\\/]node_modules[\\/]/.test(clean)) return null;
  for (let dir = dirname(clean); /[\\/]node_modules[\\/]/.test(dir + "/"); dir = dirname(dir)) {
    if (readPackage(dir)) return dir;
    if (dirname(dir) === dir) break;
  }
  return null;
}

function licenseName(pkg) {
  if (typeof pkg.license === "string") return pkg.license;
  if (pkg.license?.type) return String(pkg.license.type);
  if (Array.isArray(pkg.licenses)) return pkg.licenses.map((l) => l.type ?? l).join(" OR ");
  return "ukendt";
}

function repository(pkg) {
  const repo = typeof pkg.repository === "string" ? pkg.repository : pkg.repository?.url;
  const url = repo ?? pkg.homepage;
  return typeof url === "string" ? url.replace(/^git\+/, "").replace(/\.git$/, "") : undefined;
}

function author(pkg) {
  const a = pkg.author;
  const name = typeof a === "string" ? a : a?.name;
  // Kun navnet: e-mail og hjemmeside i feltet hører ikke hjemme i en licensoversigt.
  return typeof name === "string" ? name.replace(/\s*[<(].*$/, "").trim() || undefined : undefined;
}

/** Oplysninger og licenstekst for pakken i en mappe. `text` er tom, hvis pakken ingen licensfil har. */
export function describePackage(dir) {
  const pkg = readPackage(dir);
  if (!pkg) return null;
  const files = readdirSync(dir)
    .filter((f) => LICENSE_FILE.test(f))
    .sort();
  const text = files
    .map((f) => readFileSync(join(dir, f), "utf8").replace(/\r\n/g, "\n").trim())
    .filter(Boolean)
    .join("\n\n");
  return { name: pkg.name, version: pkg.version, license: licenseName(pkg), author: author(pkg), repository: repository(pkg), text };
}

/** Pakkerne i en række mapper: uden dubletter og sorteret, så filen er ens fra byg til byg. */
export function collectPackages(dirs, { exclude = (name) => name.startsWith("@miai/") } = {}) {
  const seen = new Map();
  for (const dir of dirs) {
    const info = describePackage(dir);
    if (info && !exclude(info.name)) seen.set(`${info.name}@${info.version}`, info);
  }
  return [...seen.values()].sort((a, b) => a.name.localeCompare(b.name, "en") || a.version.localeCompare(b.version, "en"));
}

/** Finder en afhængighed, som Node ville: i node_modules ved pakken selv og derefter opad. */
function resolveDependency(name, from) {
  for (let dir = from; ; dir = dirname(dir)) {
    const candidate = join(dir, "node_modules", name);
    if (existsSync(join(candidate, "package.json"))) return candidate;
    if (dirname(dir) === dir) return null;
  }
}

/**
 * Alt, en pakke afhænger af ved kørsel: dependencies og optionalDependencies, transitivt.
 * devDependencies er ikke med – de udleveres ikke. Valgfrie pakker til andre platforme mangler på disken og springes over.
 */
export function runtimeClosure(startDir) {
  const out = new Set();
  const queue = [resolve(startDir)];
  while (queue.length) {
    const dir = queue.pop();
    if (out.has(dir)) continue;
    out.add(dir);
    const pkg = readPackage(dir);
    for (const name of Object.keys({ ...pkg?.dependencies, ...pkg?.optionalDependencies })) {
      const dep = resolveDependency(name, dir);
      if (dep) queue.push(dep);
    }
  }
  out.delete(resolve(startDir));
  return [...out];
}

const RULE = "=".repeat(78);

/** Selve filen. `intro` fortæller, hvad pakkerne følger med (add-in'et eller appen). */
export function renderNotices(packages, intro) {
  const head = [
    "TREDJEPARTSLICENSER / THIRD-PARTY NOTICES",
    "",
    intro,
    "MIAI selv er udgivet under MIT-licensen (se LICENSE.txt). Pakkerne nedenfor tilhører deres respektive",
    "rettighedshavere og er udgivet under de licenser, der står ved hver af dem.",
    "",
    `Antal pakker: ${packages.length}`,
  ];
  const body = packages.map((p) =>
    [
      RULE,
      `${p.name} ${p.version}`,
      `Licens: ${p.license}`,
      ...(p.repository ? [`Kilde: ${p.repository}`] : []),
      RULE,
      "",
      p.text || `(Pakken indeholder ingen licensfil. Licens ifølge pakken: ${p.license}.${p.author ? ` Ophavsret: ${p.author}.` : ""} Se kilden ovenfor.)`,
      "",
    ].join("\n"),
  );
  return [...head, "", ...body].join("\n");
}

// --- Kommandolinje -----------------------------------------------------------------------------
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [target, out] = process.argv.slice(2);
  if (target !== "app" || !out) {
    console.error("Brug: node scripts/third-party-notices.mjs app <ud-fil>");
    process.exit(1);
  }
  const root = fileURLToPath(new URL("..", import.meta.url));
  // Electron er en devDependency (den pakkes af electron-builder), men udleveres som selve appen.
  const dirs = [...runtimeClosure(join(root, "packages", "app")), join(root, "node_modules", "electron")];
  const packages = collectPackages(dirs);
  if (packages.length < 10) throw new Error(`Kun ${packages.length} pakker fundet – er afhængighederne installeret (npm ci)?`);
  const intro = [
    "MIAI App udleveres sammen med de tredjepartspakker, der er listet nedenfor. Add-in'ets egne pakker",
    "står i addin/THIRD-PARTY-NOTICES.txt. Electron indeholder desuden Chromium og Node.js; deres licenser",
    "følger med appen i LICENSES.chromium.html og LICENSE.electron.txt.",
    "",
  ].join("\n");
  mkdirSync(dirname(resolve(out)), { recursive: true });
  writeFileSync(out, renderNotices(packages, intro));
  const missing = packages.filter((p) => !p.text).map((p) => p.name);
  console.log(`${out}: ${packages.length} pakker${missing.length ? `, uden licensfil: ${missing.join(", ")}` : ""}`);
}
