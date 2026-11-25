// Genererer add-in-manifester ud fra packages/addin/manifest.template.xml (Word, Excel,
// PowerPoint) og manifest.outlook.template.xml (Outlook, MailApp – #7).
//
//   node scripts/make-manifest.mjs                            → bridge + dev, begge apps (committes)
//   node scripts/make-manifest.mjs hosted <url> <ud> [outlook] → hostet variant (fx GitHub Pages eller
//                                                              organisationens egen server)
//
// Hver variant har sit eget Id, så de kan sideloades side om side.
import { readFileSync, writeFileSync } from "node:fs";

const root = new URL("../", import.meta.url);
const TEMPLATES = {
  office: readFileSync(new URL("packages/addin/manifest.template.xml", root), "utf8"),
  outlook: readFileSync(new URL("packages/addin/manifest.outlook.template.xml", root), "utf8"),
};
const version = JSON.parse(readFileSync(new URL("package.json", root), "utf8")).version;

/** Outlook-manifesterne har egne Id'er (et MailApp-manifest kan ikke dele Id med et TaskPaneApp). */
const OUTLOOK_IDS = {
  bridge: "4e1b7c90-2d6a-4f3e-9a58-c7e2b1d0f6a4",
  dev: "a7d3e5f1-8c2b-4b9e-a061-3f5d9c8e2b17",
  hosted: "e2c9a4b6-1f7d-4e30-b8a5-6d0f2c3e9a71",
};

// functions: Excels regnearksfunktioner =MIAI.ASK() og =MIAI.MODELS() (#80, ADR-0018). De kræver
// shared runtime i Excel og er kun slået til i dev-varianten, indtil manifestet er afprøvet i rigtig
// Excel på Windows og Mac. Et manifest med en skemafejl indlæses slet ikke – heller ikke i Word og
// PowerPoint – så de udgivne varianter (bridge, hosted) er uændrede.
export const VARIANTS = {
  bridge: { id: "fdac46a3-f2b8-4e84-b704-78a669531789", base: "https://localhost:44321/", suffix: " (bridge)", label: "serveres af MIAI Bridge" },
  dev: {
    id: "3b0f5a2e-7c41-4d8a-9e55-2f6c1d9b8a10",
    base: "https://localhost:3000/",
    suffix: " (dev)",
    label: "Vite-udviklingsserver, direkte forbindelse",
    functions: true,
  },
  hosted: { id: "9c7e2d14-5a3b-4f6e-8d21-b4a9e0c7f352", base: null, suffix: "", label: "statisk hostet, direkte forbindelse" },
};

/**
 * Det, regnearksfunktionerne tilføjer til manifestet. Uden dem er pladsholderne tomme, og Excel
 * åbner task pane som før (med TaskpaneId; med shared runtime må den ikke være der).
 * Der sættes bevidst intet SharedRuntime-krav øverst i manifestet: det ville skjule MIAI i Word og
 * PowerPoint på klienter uden shared runtime.
 */
const FUNCTION_PARTS = {
  EXCEL_FUNCTIONS: `
          <Runtimes>
            <Runtime resid="MIAI.Taskpane.Url" lifetime="long"/>
          </Runtimes>
          <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
              <Script>
                <SourceLocation resid="MIAI.Functions.Script.Url"/>
              </Script>
              <Page>
                <SourceLocation resid="MIAI.Taskpane.Url"/>
              </Page>
              <Metadata>
                <SourceLocation resid="MIAI.Functions.Metadata.Url"/>
              </Metadata>
              <Namespace resid="MIAI.Functions.Namespace"/>
            </ExtensionPoint>
          </AllFormFactors>`,
  EXCEL_TASKPANE_ID: "",
  FUNCTION_URLS: `
        <bt:Url id="MIAI.Functions.Script.Url" DefaultValue="{{BASE}}functions-fallback.js"/>
        <bt:Url id="MIAI.Functions.Metadata.Url" DefaultValue="{{BASE}}functions.json">
          <bt:Override Locale="en-US" Value="{{BASE}}functions.en.json"/>
        </bt:Url>`,
  FUNCTION_STRINGS: `
        <bt:String id="MIAI.Functions.Namespace" DefaultValue="MIAI"/>`,
};
const NO_FUNCTION_PARTS = {
  EXCEL_FUNCTIONS: "",
  EXCEL_TASKPANE_ID: `
                      <TaskpaneId>MIAI.Taskpane</TaskpaneId>`,
  FUNCTION_URLS: "",
  FUNCTION_STRINGS: "",
};

function render(variant, base, app = "office") {
  const v = { ...VARIANTS[variant], ...(app === "outlook" ? { id: OUTLOOK_IDS[variant] } : {}) };
  const b = (base ?? v.base).replace(/\/?$/, "/");
  const origin = new URL(b).origin;
  if (!b.startsWith("https://")) throw new Error(`Office kræver HTTPS: ${b}`);
  let xml = TEMPLATES[app];
  for (const [key, value] of Object.entries(v.functions && app === "office" ? FUNCTION_PARTS : NO_FUNCTION_PARTS)) xml = xml.replaceAll(`{{${key}}}`, value);
  return xml
    .replaceAll("{{BASE}}", b)
    .replaceAll("{{ORIGIN}}", origin)
    .replaceAll("{{ID}}", v.id)
    .replaceAll("{{VERSION}}", `${version}.0`)
    .replaceAll("{{SUFFIX}}", v.suffix)
    .replaceAll("{{VARIANT}}", `${variant}: ${v.label} (${b})`);
}

const [variant, base, out, app = "office"] = process.argv.slice(2);
if (variant) {
  writeFileSync(out, render(variant, base, app));
  console.log(`✓ ${variant} (${app}) → ${out}`);
} else {
  for (const [name, file, kind] of [
    ["bridge", "manifest.xml", "office"],
    ["dev", "manifest.dev.xml", "office"],
    ["bridge", "manifest.outlook.xml", "outlook"],
    ["dev", "manifest.outlook.dev.xml", "outlook"],
  ]) {
    writeFileSync(new URL(`packages/addin/${file}`, root), render(name, undefined, kind));
    console.log(`✓ ${name} (${kind}) → packages/addin/${file}`);
  }
}
