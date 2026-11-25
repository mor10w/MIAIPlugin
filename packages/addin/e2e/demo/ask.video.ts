import { readFileSync } from "node:fs";
import { createServer, type Server } from "node:http";
import { expect, test, type FrameLocator, type Page } from "@playwright/test";
import { fakeOfficeScript, type FakeOfficeOptions } from "../fixtures/office";
import { startBridge } from "../fixtures/servers";

/**
 * Video om regnearksfunktionen =MIAI.ASK() (#80), til brug i opslag: formlen skrives i en celle,
 * trækkes ned over rækkerne, svaret følger med, når teksten ændres, og til sidst slår en celle op
 * på nettet (optimized = FALSK).
 *
 * Regnearket er scenens eget (stage.html, "formelscenen"): den spiller Excels rolle, viser #BUSY!
 * og kalder den funktion, add-in'et har registreret i task pane. Task pane serveres af den rigtige
 * bridge, som i MIAI App, så køen, cachen, statuslinjen "MIAI.ASK: … i gang, … i kø" og
 * søgerunden (modellen beder om web_search, bridgen søger og sender resultatet tilbage) er MIAI's
 * egen kode. Modelsvarene og søgeresultaterne er indspillede (ROWS og COMPANIES nedenfor), så
 * optagelsen er ens fra gang til gang og ikke rører internettet. Det er en illustration af
 * brugerfladen, ikke en måling af en models kvalitet eller hastighed.
 *
 *   npm run build && npx playwright test -c packages/addin/e2e/demo/playwright.config.ts ask.video.ts
 *
 * Optagelsen ender i packages/addin/demo-results/…/video.webm. Til MP4, se docs/UDVIKLING.md.
 */

const PROBE = !!process.env.MIAI_DEMO_PROBE;

// --- Regnearket og de indspillede svar --------------------------------------------------------
/** Henvendelse, og det svar modellen giver på hver af de to prompter. */
const ROWS: [text: string, category: string, mood: string][] = [
  ["Min pakke er ikke kommet endnu, bestilt i mandags", "Levering", "Negativ"],
  ["Fakturaen er trukket to gange på mit kort", "Faktura", "Negativ"],
  ["Tak for hurtig hjælp med opsætningen i går!", "Support", "Positiv"],
  ["Kan I sende en kopi af fakturaen for marts?", "Faktura", "Neutral"],
  ["Appen lukker ned, når jeg åbner en rapport", "Support", "Negativ"],
  ["Hvornår kan I levere til Bornholm?", "Levering", "Neutral"],
  ["Super service, pakken kom allerede dagen efter", "Levering", "Positiv"],
  ["Jeg kan ikke logge ind efter opdateringen", "Support", "Negativ"],
];
/** Rækken, der rettes i sidste scene (række 7 i arket), og dens nye svar. */
const EDIT_CELL = "A7";
const EDITED: (typeof ROWS)[number] = ["Jeg har fået en rykker, men har betalt fakturaen", "Faktura", "Negativ"];

const FIRST = 2;
const LAST = FIRST + ROWS.length - 1;
const CATEGORY = (row: number) => `=MIAI.ASK("Kategorisér som Faktura, Levering eller Support"; A${row}; SAND)`;
const MOOD = (row: number) => `=MIAI.ASK("Er tonen Positiv, Neutral eller Negativ?"; A${row}; SAND)`;

const SHEET: NonNullable<FakeOfficeOptions["excel"]> = {
  sheet: "Henvendelser",
  values: [["Henvendelse", "Kategori", "Stemning"], ...ROWS.map(([text]) => [text, "", ""])],
  formats: { A1: { bold: true }, B1: { bold: true }, C1: { bold: true } },
};

/** Arket til websøgningen: kendte danske virksomheder og byen med hovedkontoret. */
const COMPANIES: [name: string, city: string][] = [
  ["LEGO", "Billund"],
  ["Novo Nordisk", "Bagsværd"],
  ["Vestas", "Aarhus"],
  ["Danfoss", "Nordborg"],
  ["Grundfos", "Bjerringbro"],
  ["Carlsberg", "København"],
];
const LAST_COMPANY = FIRST + COMPANIES.length - 1;
const HEAD_OFFICE = (row: number) => `=MIAI.ASK("Hvilken by ligger hovedkontoret i?"; A${row}; FALSK)`;
const COMPANY_SHEET: NonNullable<FakeOfficeOptions["excel"]> = {
  sheet: "Leverandører",
  values: [["Virksomhed", "Hovedkontor"], ...COMPANIES.map(([name]) => [name, ""])],
  formats: { A1: { bold: true }, B1: { bold: true } },
};

/** Lille søgetjeneste (SearXNG's JSON-format), så bridgens websøgning ikke rører internettet. */
async function startSearch(): Promise<{ url: string; queries: string[]; stop: () => Promise<void> }> {
  const queries: string[] = [];
  const server: Server = createServer((req, res) => {
    const q = new URL(req.url ?? "/", "http://x").searchParams.get("q") ?? "";
    queries.push(q);
    const hit = COMPANIES.find(([name]) => q.includes(name));
    res.setHeader("content-type", "application/json");
    res.end(
      JSON.stringify({
        results: hit
          ? [{ title: `${hit[0]} – hovedkontor`, url: "https://example.invalid/virksomheder", content: `${hit[0]} har hovedkontor i ${hit[1]}.` }]
          : [],
      }),
    );
  });
  await new Promise<void>((ok) => server.listen(0, "127.0.0.1", ok));
  const { port } = server.address() as { port: number };
  return { url: `http://127.0.0.1:${port}`, queries, stop: () => new Promise((ok) => server.close(() => ok())) };
}

/**
 * Lille OpenAI-kompatibel model: finder rækken i forespørgslen og svarer det indspillede. Til en
 * virksomhed beder den først om en websøgning (web_search) og svarer, når søgeresultatet er kommet.
 */
async function startModel(): Promise<{ url: string; stop: () => Promise<void> }> {
  const sleep = (ms: number) => new Promise((ok) => setTimeout(ok, ms));
  let n = 0;
  const server: Server = createServer(async (req, res) => {
    res.setHeader("access-control-allow-origin", "*");
    res.setHeader("access-control-allow-headers", "content-type, authorization");
    if (req.method === "OPTIONS") return res.end();
    if (req.url === "/v1/models") {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ object: "list", data: [{ id: "qwen3:14b" }] }));
    }
    let raw = "";
    for await (const chunk of req) raw += chunk;
    if (req.url !== "/v1/chat/completions") {
      res.statusCode = 404;
      return res.end();
    }
    const body = JSON.parse(raw) as { messages: { role: string; content: string | null }[]; stream?: boolean };
    const asked = body.messages.map((m) => m.content ?? "").join("\n");
    const user = body.messages.find((m) => m.role === "user")?.content ?? "";
    await sleep(PROBE ? 30 : 950 + (n++ % 3) * 220); // "tænketid", lidt forskellig fra celle til celle
    const company = COMPANIES.find(([name]) => user.includes(`\n${name}\n`));
    if (company && !body.messages.some((m) => m.role === "tool")) {
      // Første runde for en virksomhed: modellen beder om en websøgning.
      const call = {
        index: 0,
        id: "call_1",
        type: "function",
        function: { name: "web_search", arguments: JSON.stringify({ query: `${company[0]} hovedkontor` }) },
      };
      if (!body.stream) {
        res.setHeader("content-type", "application/json");
        return res.end(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [call] }, finish_reason: "tool_calls" }] }));
      }
      res.writeHead(200, { "content-type": "text/event-stream" });
      res.write(`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [call] }, finish_reason: "tool_calls" }] })}\n\n`);
      return res.end("data: [DONE]\n\n");
    }
    const row = [EDITED, ...ROWS].find(([text]) => asked.includes(text));
    const answer = company ? company[1] : row ? (asked.includes("Kategorisér") ? row[1] : row[2]) : "?";
    if (!body.stream) {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ choices: [{ message: { content: answer }, finish_reason: "stop" }] }));
    }
    res.writeHead(200, { "content-type": "text/event-stream" });
    res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: answer }, finish_reason: "stop" }] })}\n\n`);
    res.end("data: [DONE]\n\n");
  });
  await new Promise<void>((ok) => server.listen(0, "127.0.0.1", ok));
  const { port } = server.address() as { port: number };
  return { url: `http://127.0.0.1:${port}/v1`, stop: () => new Promise((ok) => server.close(() => ok())) };
}

// --- Scenen -----------------------------------------------------------------------------------
const pause = (page: Page, ms: number) => page.waitForTimeout(PROBE ? 30 : ms);
type Stage = {
  open(url: string, name: string): void;
  caption(n: string, t: string, s?: string): void;
  cursor(x: number, y: number): void;
  ripple(x: number, y: number): void;
  card(title?: string, text?: string, tags?: string[], small?: string): void;
  grid(opts: { cols: string[]; widths: Record<string, number>; rows: number }): void;
  fsel(range: string): void;
  fedit(a: string, text: string): void;
  fcommit(a: string, text: string): void;
  ffill(src: string, range: string): void;
};
const stage = <K extends keyof Stage>(page: Page, fn: K, ...args: Parameters<Stage[K]>) =>
  page.evaluate(([f, a]) => (window as unknown as { stage: Record<string, (...x: unknown[]) => void> }).stage[f as string](...(a as unknown[])), [
    fn,
    args,
  ] as const);

const cell = (page: Page, a: string) => page.locator(`#sheet td[data-a="${a}"]`);

/** Et punkt i cellen: midten, eller hjørnet med fyldhåndtaget. */
async function point(page: Page, a: string, where: "center" | "handle" = "center") {
  const box = (await cell(page, a).boundingBox())!;
  return where === "handle" ? { x: box.x + box.width - 2, y: box.y + box.height - 2 } : { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

/** Flytter den tegnede markør hen til cellen og markerer den. */
async function select(page: Page, a: string) {
  const { x, y } = await point(page, a);
  await stage(page, "cursor", x, y);
  await pause(page, 750);
  await stage(page, "ripple", x, y);
  await stage(page, "fsel", a);
  await pause(page, 450);
}

/** Skriver i cellen tegn for tegn og afslutter med Enter. `from` er det, der allerede står i cellen. */
async function write(page: Page, a: string, text: string, from = "") {
  await select(page, a);
  let keep = from.length;
  // Den gamle tekst slettes baglæns, hurtigere end der skrives.
  for (; keep > 0; keep -= 3) {
    await stage(page, "fedit", a, from.slice(0, Math.max(0, keep - 3)));
    await page.waitForTimeout(PROBE ? 0 : 14);
  }
  for (let i = 1; i <= text.length; i++) {
    await stage(page, "fedit", a, text.slice(0, i));
    // Lidt længere ophold efter et skilletegn, som når man skriver en formel.
    await page.waitForTimeout(PROBE ? 0 : /[;(]/.test(text[i - 1]) ? 150 : 30);
  }
  // Prøvekørsel: et stillbillede af cellen, mens den redigeres.
  if (PROBE) await page.screenshot({ path: test.info().outputPath(`skriv-${a}.png`) });
  await pause(page, 900);
  await stage(page, "fcommit", a, text);
}

/** Trækker fyldhåndtaget fra `src` ned til række `last`; formlerne beregnes, når der slippes. */
async function fillDown(page: Page, src: string, last: number) {
  const col = src[0];
  const start = await point(page, src, "handle");
  await stage(page, "cursor", start.x, start.y);
  await pause(page, 800);
  await stage(page, "ripple", start.x, start.y);
  for (let row = Number(src.slice(1)) + 1; row <= last; row++) {
    const p = await point(page, `${col}${row}`, "handle");
    await stage(page, "cursor", p.x, p.y);
    await stage(page, "fsel", `${src}:${col}${row}`);
    await pause(page, 130);
  }
  await pause(page, 450);
  await stage(page, "ffill", src, `${src}:${col}${last}`);
}

test("video: =MIAI.ASK()", async ({ page }) => {
  if (PROBE) page.on("pageerror", (e) => console.log("[scene]", e.message));
  const model = await startModel();
  const search = await startSearch();
  // Den rigtige bridge, som i MIAI App: websøgning i en celle findes kun dér.
  const bridge = await startBridge(model.url, { defaultModel: "qwen3:14b", webSearch: true, search: { provider: "searxng", url: search.url } });
  const pane = (sheet: string) => `${bridge.origin}/taskpane.html?token=${encodeURIComponent(bridge.token)}&functions=1&ark=${sheet}`;
  let office: FakeOfficeOptions = {
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: SHEET,
    customFunctions: true,
    documentUrl: "https://example.invalid/Henvendelser.xlsx",
  };
  await page.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: fakeOfficeScript(office) }));
  await page.route(`${bridge.origin}/demo-stage.html`, (r) =>
    r.fulfill({ contentType: "text/html", body: readFileSync(new URL("./stage.html", import.meta.url), "utf8") }),
  );

  await page.goto(`${bridge.origin}/demo-stage.html`);
  await stage(
    page,
    "card",
    "=MIAI.ASK()",
    "Brug jeres egen sprogmodel som en formel i Excel.",
    ["Modellen kører lokalt", "Svaret står i cellen", "Træk formlen ned"],
    "Demo med simuleret regneark og indspillede modelsvar",
  );
  // Et manifest med regnearksfunktionerne åbner task pane med ?functions=1.
  await stage(page, "open", pane("henvendelser"), "Henvendelser.xlsx");
  const tp: FrameLocator = page.frameLocator("#tp");
  await expect(tp.getByText("Via MIAI Bridge")).toBeVisible();
  await expect(page.locator("#sheet td", { hasText: "Bornholm" })).toBeVisible();
  await stage(page, "grid", { cols: ["A", "B", "C", "D", "E"], widths: { A: 350, B: 130, C: 120, D: 110, E: 110 }, rows: 14 });
  const bar = tp.getByTestId("cell-bar");

  /** Prøvekørsel (MIAI_DEMO_PROBE=1): hurtig gennemgang med et stillbillede pr. scene i stedet for pauser. */
  const snap = async (name: string) => {
    if (PROBE) await page.screenshot({ path: test.info().outputPath(`${name}.png`) });
  };

  await pause(page, 3600);
  await stage(page, "card");
  await pause(page, 900);

  // --- 1. Formlen i én celle -------------------------------------------------------------------
  await stage(page, "caption", "1 · Formlen", "Skriv prompten som en formel", "Prompt, cellen med teksten og SAND for et kort, hurtigt svar.");
  await pause(page, 1500);
  await write(page, `B${FIRST}`, CATEGORY(FIRST));
  await expect(cell(page, `B${FIRST}`)).toHaveText("#BUSY!");
  await snap("1-busy");
  await expect(cell(page, `B${FIRST}`)).toHaveText(ROWS[0][1]);
  await stage(page, "fsel", `B${FIRST}`);
  await pause(page, 2400);
  await snap("1-svar");

  // --- 2. Træk formlen ned ---------------------------------------------------------------------
  await stage(page, "caption", "2 · Træk ned", "Samme formel på alle rækker", "MIAI sender to celler ad gangen til modellen. Køen ses i panelet.");
  await pause(page, 1300);
  await fillDown(page, `B${FIRST}`, LAST);
  await expect(bar).toContainText("2 i gang");
  await snap("2-koe");
  for (const [i, row] of ROWS.entries()) await expect(cell(page, `B${FIRST + i}`)).toHaveText(row[1]);
  await expect(bar).toBeHidden();
  await pause(page, 2400);
  await snap("2-faerdig");

  // --- 3. En ny prompt i næste kolonne ---------------------------------------------------------
  await stage(page, "caption", "3 · Ny kolonne", "En ny prompt giver en ny kolonne", "Svaret står som værdi i cellen, klar til filtre og pivottabeller.");
  await pause(page, 1300);
  await write(page, `C${FIRST}`, MOOD(FIRST));
  await expect(cell(page, `C${FIRST}`)).toHaveText(ROWS[0][2]);
  await stage(page, "fsel", `C${FIRST}`);
  await pause(page, 700);
  await fillDown(page, `C${FIRST}`, LAST);
  for (const [i, row] of ROWS.entries()) await expect(cell(page, `C${FIRST + i}`)).toHaveText(row[2]);
  await expect(bar).toBeHidden();
  await pause(page, 2400);
  await snap("3-faerdig");

  // --- 4. Ret teksten: cellerne beregnes igen --------------------------------------------------
  await stage(
    page,
    "caption",
    "4 · Ret teksten",
    "Svaret følger med, når teksten ændres",
    "Kun den ændrede række spørger modellen igen. Resten ligger i cachen.",
  );
  await pause(page, 1300);
  const row = Number(EDIT_CELL.slice(1));
  await write(page, EDIT_CELL, EDITED[0], ROWS[row - FIRST][0]);
  await expect(cell(page, `B${row}`)).toHaveText("#BUSY!");
  await snap("4-busy");
  await expect(cell(page, `B${row}`)).toHaveText(EDITED[1]);
  await expect(cell(page, `C${row}`)).toHaveText(EDITED[2]);
  await stage(page, "fsel", `B${row}:C${row}`);
  await pause(page, 3400);
  await snap("4-faerdig");

  // --- 5. Websøgning i cellen (optimized = FALSK) ----------------------------------------------
  await stage(page, "card", "Og med websøgning", "Med FALSK som sidste argument må cellen bruge de værktøjer, du har slået til under Indstillinger.", [
    "Søger på nettet",
    "Egne mapper",
    "Tænkning",
  ]);
  await pause(page, 900);
  office = {
    host: "Excel",
    apiSets: { ExcelApi: "1.10" },
    excel: COMPANY_SHEET,
    customFunctions: true,
    documentUrl: "https://example.invalid/Leverandører.xlsx",
  };
  await stage(page, "open", pane("leverandoerer"), "Leverandører.xlsx");
  await expect(tp.getByText("Via MIAI Bridge")).toBeVisible();
  await expect(page.locator("#sheet td", { hasText: "Grundfos" })).toBeVisible();
  await stage(page, "grid", { cols: ["A", "B", "C", "D", "E"], widths: { A: 220, B: 200, C: 140, D: 130, E: 130 }, rows: 14 });
  await stage(
    page,
    "caption",
    "5 · Websøgning",
    "FALSK: cellen må slå op på nettet",
    "Modellen søger, læser resultatet og svarer. Det tager længere tid end SAND.",
  );
  await pause(page, 2600);
  await stage(page, "card");
  await pause(page, 1800);
  await write(page, `B${FIRST}`, HEAD_OFFICE(FIRST));
  await expect(cell(page, `B${FIRST}`)).toHaveText("#BUSY!");
  await expect(cell(page, `B${FIRST}`)).toHaveText(COMPANIES[0][1]);
  await stage(page, "fsel", `B${FIRST}`);
  await pause(page, 1600);
  await fillDown(page, `B${FIRST}`, LAST_COMPANY);
  await expect(bar).toContainText("2 i gang");
  await snap("5-koe");
  for (const [i, company] of COMPANIES.entries()) await expect(cell(page, `B${FIRST + i}`)).toHaveText(company[1]);
  await expect(bar).toBeHidden();
  // Hver virksomhed er slået op gennem bridgens websøgning, ikke svaret direkte.
  for (const [name] of COMPANIES) expect(search.queries).toContain(`${name} hovedkontor`);
  await pause(page, 3200);
  await snap("5-faerdig");

  // --- Afslutning ------------------------------------------------------------------------------
  await stage(
    page,
    "card",
    "=MIAI.ASK(prompt; input)",
    "En del af MIAI: AI-assistent til Word, Excel, PowerPoint og Outlook med jeres egne sprogmodeller. Åben kildekode under MIT-licensen.",
    ["Kategorisér", "Udtræk", "Oversæt", "Slå op på nettet"],
    "github.com/mor10w/MIAIPlugin",
  );
  await pause(page, 5000);

  await page.close();
  await bridge.stop();
  await search.stop();
  await model.stop();
});
