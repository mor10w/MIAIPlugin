// Genererer add-in-manifester ud fra packages/addin/manifest.template.xml.
//
//   node scripts/make-manifest.mjs                  → bridge + dev (committes)
//   node scripts/make-manifest.mjs hosted <url> <ud> → hostet variant (fx GitHub Pages, bruges i CI)
//
// Hver variant har sit eget Id, så de kan sideloades side om side.
import { readFileSync, writeFileSync } from "node:fs";

const root = new URL("../", import.meta.url);
const template = readFileSync(new URL("packages/addin/manifest.template.xml", root), "utf8");
const version = JSON.parse(readFileSync(new URL("package.json", root), "utf8")).version;

export const VARIANTS = {
  bridge: { id: "fdac46a3-f2b8-4e84-b704-78a669531789", base: "https://localhost:44321/", suffix: " (bridge)", label: "serveres af MIAI Bridge" },
  dev: { id: "3b0f5a2e-7c41-4d8a-9e55-2f6c1d9b8a10", base: "https://localhost:3000/", suffix: " (dev)", label: "Vite-udviklingsserver, direkte forbindelse" },
  hosted: { id: "9c7e2d14-5a3b-4f6e-8d21-b4a9e0c7f352", base: null, suffix: "", label: "statisk hostet, direkte forbindelse" },
};

function render(variant, base) {
  const v = VARIANTS[variant];
  const b = (base ?? v.base).replace(/\/?$/, "/");
  const origin = new URL(b).origin;
  if (!b.startsWith("https://")) throw new Error(`Office kræver HTTPS: ${b}`);
  return template
    .replaceAll("{{BASE}}", b)
    .replaceAll("{{ORIGIN}}", origin)
    .replaceAll("{{ID}}", v.id)
    .replaceAll("{{VERSION}}", `${version}.0`)
    .replaceAll("{{SUFFIX}}", v.suffix)
    .replaceAll("{{VARIANT}}", `${variant}: ${v.label} (${b})`);
}

const [variant, base, out] = process.argv.slice(2);
if (variant) {
  writeFileSync(out, render(variant, base));
  console.log(`✓ ${variant} → ${out}`);
} else {
  for (const [name, file] of [["bridge", "manifest.xml"], ["dev", "manifest.dev.xml"]]) {
    writeFileSync(new URL(`packages/addin/${file}`, root), render(name));
    console.log(`✓ ${name} → packages/addin/${file}`);
  }
}
