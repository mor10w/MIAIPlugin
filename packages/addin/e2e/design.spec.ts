import AxeBuilder from "@axe-core/playwright";
import JSZip from "jszip";
import { expect, test } from "./fixtures/test";

/** PowerPoint: design-master ud fra indholdet (ADR-0014). */

const SLIDES = [
  {
    layout: "Title Slide",
    shapes: [
      { name: "Title 1", ph: "Title", text: "Kvartalsrapport" },
      { name: "Subtitle 2", ph: "Subtitle", text: "Q3 2026" },
    ],
  },
  {
    layout: "Title and Content",
    shapes: [
      { name: "Title 1", ph: "Title", text: "Resultater" },
      { name: "Content 2", ph: "Body", text: "Omsætning +12 %\nMargin 38 %" },
    ],
  },
  {
    layout: "Title and Content",
    shapes: [
      { name: "Title 1", ph: "Title", text: "Næste skridt" },
      { name: "Content 2", ph: "Body", text: "Webshop i Sverige" },
    ],
  },
];

const PLAN = {
  name: "Nordisk klar",
  rationale: "Rolig og troværdig til bestyrelsen; blå signalerer stabilitet, orange fremhæver tallene.",
  style: "clean",
  palette: { background: "#FFFFFF", text: "#1B2A3A", muted: "#5B6B7B", primary: "#0A5C8A", accent: "#E07A1F" },
  fonts: { heading: "Georgia", body: "Aptos" },
  slides: [
    { layout: "title", title: "Kvartalsrapport Q3", subtitle: "Bestyrelsesmøde" },
    {
      layout: "kpi",
      title: "Resultater",
      kpis: [
        { value: "+12 %", label: "Omsætning" },
        { value: "38 %", label: "Margin" },
      ],
    },
    { layout: "section", title: "Næste skridt" },
    { layout: "bullets", title: "Fokus", bullets: ["Webshop i Sverige", "To nye butikker"] },
    { layout: "closing", title: "Tak", subtitle: "Spørgsmål?" },
  ],
};

type Deck = { layout: string; shapes: { text: string }[] }[];
const deck = (page: import("@playwright/test").Page) => page.evaluate(() => (window as unknown as { __miai: { deck: () => Deck } }).__miai.deck());

test("design: forhåndsvisning af palet og slides; masteren sættes ind, de gamle slides erstattes og kommer tilbage med Fortryd", async ({ open }) => {
  const tp = await open({ host: "PowerPoint", apiSets: { PowerPointApi: "1.8" }, slides: SLIDES });
  tp.model.reply = () => ({ json: JSON.stringify(PLAN) });
  await tp.run("Design: master og layout");
  await expect(tp.page.getByText("Nordisk klar", { exact: true })).toBeVisible();

  // Modellen fik indholdet og designreglerne.
  const msgs = tp.model.lastRequest!.messages;
  expect(msgs[1].content).toContain("2.2 [body] Content 2: Omsætning +12 % ⏎ Margin 38 %");
  expect(msgs[0].content).toMatch(/ONE consistent look/);

  // Forhåndsvisning: palet med hex og en miniature pr. slide.
  await expect(tp.page.getByText("Primær #0A5C8A")).toBeVisible();
  for (const name of ["1. Titel", "2. Nøgletal", "3. Sektion", "4. Punkter", "5. Afslutning"]) await expect(tp.page.getByRole("img", { name })).toBeVisible();
  const axe = await new AxeBuilder({ page: tp.page }).withTags(["wcag2a", "wcag2aa", "wcag21a", "wcag21aa"]).exclude("[data-tabster-dummy]").analyze();
  expect(axe.violations.filter((v) => v.impact === "serious" || v.impact === "critical").map((v) => v.id)).toEqual([]);

  await tp.page.getByRole("switch", { name: "Erstat de nuværende 3 slides" }).check();
  await tp.page.getByRole("button", { name: "Opret 5 slides med designet" }).click();
  await expect(tp.page.getByText("5 slides oprettet med designet. 3 gamle slides er fjernet.")).toBeVisible();
  expect((await deck(tp.page)).map((s) => s.layout)).toEqual(Array(5).fill("(fra designet)"));

  // Filen, der blev sat ind, har designets master og layouts – og den blev sat ind med kildeformatering.
  expect(await tp.find("insertPptx")).toMatchObject({ formatting: "KeepSourceFormatting", slides: 5 });
  const b64 = await tp.page.evaluate(() => (window as unknown as { __miai: { pptx: string[] } }).__miai.pptx[0]);
  const zip = await JSZip.loadAsync(Buffer.from(b64, "base64"));
  const layouts = await Promise.all(
    Object.keys(zip.files)
      .filter((f) => /slideLayouts\/slideLayout\d+\.xml$/.test(f))
      .map((f) => zip.file(f)!.async("string")),
  );
  const names = layouts.map((x) => x.match(/<p:cSld name="([^"]+)"/)?.[1]);
  expect(names).toEqual(expect.arrayContaining(["Nordisk klar – Titel", "Nordisk klar – Nøgletal", "Nordisk klar – Punkter", "Nordisk klar – Afslutning"]));

  // Fortryd: designets slides fjernes, og de gamle kommer tilbage i rækkefølge.
  await tp.page.getByRole("button", { name: "Fortryd" }).click();
  await expect(tp.page.getByText("Fortrudt. Præsentationen er som før.")).toBeVisible();
  expect((await deck(tp.page)).map((s) => s.shapes[0].text)).toEqual(["Kvartalsrapport", "Resultater", "Næste skridt"]);
});

test("design i ældre PowerPoint: de gamle slides bevares (kan ikke fortrydes), og brugeren får besked", async ({ open }) => {
  const tp = await open({ host: "PowerPoint", apiSets: { PowerPointApi: "1.4" }, slides: SLIDES });
  tp.model.reply = () => ({ json: JSON.stringify(PLAN) });
  await tp.run("Design: master og layout");
  await tp.page.getByRole("switch", { name: "Erstat de nuværende 3 slides" }).check();
  await tp.page.getByRole("button", { name: "Opret 5 slides med designet" }).click();
  await expect(tp.page.getByText(/De gamle slides er bevaret/)).toBeVisible();
  expect((await deck(tp.page)).length).toBe(8);
});
