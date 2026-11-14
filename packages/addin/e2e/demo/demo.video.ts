import { readFileSync } from "node:fs";
import { createServer, type Server } from "node:http";
import { expect, test, type FrameLocator, type Locator, type Page } from "@playwright/test";
import { fakeOfficeScript, type FakeOfficeOptions } from "../fixtures/office";
import { startStaticHost } from "../fixtures/servers";

/**
 * Demovideo af MIAI: task pane ved siden af et simuleret dokument, styret trin for trin.
 * Dokumentet er den simulerede Word fra browsertestene, og modelsvarene er indspillede
 * (SCRIPT nedenfor), så optagelsen er ens fra gang til gang. Det er altså en illustration af
 * brugerfladen, ikke en måling af en models kvalitet eller hastighed.
 *
 *   npm run build && npx playwright test -c packages/addin/e2e/demo/playwright.config.ts
 *
 * Optagelsen ender i packages/addin/demo-results/…/video.webm. Til MP4:
 *   ffmpeg -i video.webm -c:v libx264 -crf 20 -pix_fmt yuv420p -movflags +faststart -an miai-demo.mp4
 */

const PROBE = !!process.env.MIAI_DEMO_PROBE;

// --- Dokumentet -----------------------------------------------------------------------------
const DOC: NonNullable<FakeOfficeOptions["paragraphs"]> = [
  { text: "Samarbejdsaftale om drift af dataplatform", style: "Heading1" },
  {
    text: "Denne aftale beskriver samarbejdet mellem Leverandøren og Kunden om drift af Kundens dataplatform. Aftalen træder i kraft d. 1 januar 2027 og løber i 24 måneder.",
  },
  { text: "Leverandøren leverer en løsning der opfylder kravene i bilag 1. Kunden skal godkende leverancen senest 30 dage efter modtagelsen." },
  { text: "Betaling sker månedsvis bagud med en frist på 14 dage. Ved forsinket betaling tilskrives rente efter rentelovens bestemelser." },
  { text: "Parterne behandler alle oplysninger fortroligt, også efter aftalens ophør. Aftalen kan opsiges med 6 måneders varsel." },
];

// --- Indspillede modelsvar --------------------------------------------------------------------
type Reply = { json: string } | { stream: string };
const SCRIPT: Record<"grammar" | "agent" | "chat", Reply> = {
  grammar: {
    json: JSON.stringify({
      edits: [
        { original: "d. 1 januar", replacement: "d. 1. januar", explanation: "Punktum efter ordenstal" },
        { original: "en løsning der opfylder", replacement: "en løsning, der opfylder", explanation: "Komma før ledsætning" },
        { original: "bestemelser", replacement: "bestemmelser", explanation: "Stavefejl" },
      ],
    }),
  },
  agent: {
    json: JSON.stringify({
      summary: "Giver afsnittene overskrifter og samler aftalens frister i en tabel til sidst.",
      actions: [
        { action: "insert", position: "before", paragraph: 2, text: "Formål og løbetid", style: "heading2", reason: "Overskrift til indledningen" },
        { action: "insert", position: "before", paragraph: 3, text: "Leverance", style: "heading2", reason: "Overskrift" },
        { action: "insert", position: "before", paragraph: 4, text: "Betaling", style: "heading2", reason: "Overskrift" },
        { action: "insert", position: "before", paragraph: 5, text: "Fortrolighed og opsigelse", style: "heading2", reason: "Overskrift" },
        {
          action: "insertTable",
          position: "end",
          rows: [
            ["Frist", "Hvad", "Hvem"],
            ["30 dage", "Godkendelse af leverancen", "Kunden"],
            ["14 dage", "Betaling af faktura", "Kunden"],
            ["6 måneder", "Opsigelsesvarsel", "Begge parter"],
          ],
          reason: "Overblik over fristerne i aftalen",
        },
      ],
    }),
  },
  chat: {
    stream:
      "Kunden har to forpligtelser med frist:\n\n- **Godkende leverancen** senest 30 dage efter modtagelsen.\n- **Betale** månedsvis bagud inden for 14 dage; ellers tilskrives rente.\n\nDerudover gælder fortroligheden også efter aftalens ophør.",
  },
};

/** Lille OpenAI-kompatibel model, der svarer med de indspillede svar i et roligt tempo. */
async function startModel(pick: () => Reply): Promise<{ url: string; stop: () => Promise<void> }> {
  const sleep = (ms: number) => new Promise((ok) => setTimeout(ok, ms));
  const server: Server = createServer(async (req, res) => {
    res.setHeader("access-control-allow-origin", "*");
    res.setHeader("access-control-allow-headers", "content-type, authorization");
    if (req.method === "OPTIONS") return res.end();
    if (req.url === "/v1/models") {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ object: "list", data: [{ id: "qwen3:14b" }] }));
    }
    for await (const _ of req) void _;
    const r = pick();
    await sleep(1300); // "tænketid"
    if ("json" in r) {
      res.setHeader("content-type", "application/json");
      return res.end(JSON.stringify({ choices: [{ message: { content: r.json }, finish_reason: "stop" }] }));
    }
    res.writeHead(200, { "content-type": "text/event-stream" });
    const parts = r.stream.match(/\S+\s*/g) ?? [];
    for (const [i, content] of parts.entries()) {
      res.write(`data: ${JSON.stringify({ choices: [{ delta: { content }, finish_reason: i === parts.length - 1 ? "stop" : null }] })}\n\n`);
      await sleep(55);
    }
    res.end("data: [DONE]\n\n");
  });
  await new Promise<void>((ok) => server.listen(0, "127.0.0.1", ok));
  const { port } = server.address() as { port: number };
  return { url: `http://127.0.0.1:${port}/v1`, stop: () => new Promise((ok) => server.close(() => ok())) };
}

// --- Scenen -----------------------------------------------------------------------------------
const pause = (page: Page, ms: number) => page.waitForTimeout(PROBE ? 50 : ms);
type Stage = {
  open(url: string, name: string): void;
  caption(n: string, t: string, s?: string): void;
  track(on: boolean): void;
  select(texts: string[]): void;
  cursor(x: number, y: number): void;
  ripple(x: number, y: number): void;
  card(title?: string, text?: string, tags?: string[], small?: string): void;
};
const stage = <K extends keyof Stage>(page: Page, fn: K, ...args: Parameters<Stage[K]>) =>
  page.evaluate(([f, a]) => (window as unknown as { stage: Record<string, (...x: unknown[]) => void> }).stage[f as string](...(a as unknown[])), [
    fn,
    args,
  ] as const);

/** Flytter den tegnede markør hen til et element og klikker (optagelsen viser ikke den rigtige mus). */
async function click(page: Page, target: Locator) {
  await target.scrollIntoViewIfNeeded();
  const box = (await target.boundingBox())!;
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await stage(page, "cursor", x, y);
  await pause(page, 750);
  await stage(page, "ripple", x, y);
  await target.click();
  await pause(page, 350);
}

async function type(page: Page, field: Locator, text: string) {
  await click(page, field);
  await field.pressSequentially(text, { delay: PROBE ? 0 : 32 });
  await pause(page, 500);
}

test("demovideo", async ({ page }) => {
  let next: Reply = SCRIPT.grammar;
  const model = await startModel(() => next);
  const host = await startStaticHost();
  const office: FakeOfficeOptions = {
    host: "Word",
    apiSets: { WordApi: "1.6" },
    paragraphs: DOC,
    documentUrl: "https://example.invalid/Samarbejdsaftale.docx",
  };
  await page.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: fakeOfficeScript(office) }));
  await page.route(`${host.origin}/demo-stage.html`, (r) =>
    r.fulfill({ contentType: "text/html", body: readFileSync(new URL("./stage.html", import.meta.url), "utf8") }),
  );
  const settings = { endpoints: [{ id: "lokal", name: "Ollama (lokal)", baseUrl: model.url }], activeEndpoint: "lokal" };
  await page.addInitScript((s) => localStorage.setItem("miai.settings.v1", s), JSON.stringify(settings));

  await page.goto(`${host.origin}/demo-stage.html`);
  await stage(
    page,
    "card",
    "MIAI",
    "AI-assistent til Word, Excel, PowerPoint og Outlook – med jeres egne sprogmodeller.",
    ["Modellen kører lokalt", "Intet ændres uden godkendelse", "Alt kan fortrydes"],
    "Demo med simuleret dokument og indspillede modelsvar",
  );
  await stage(page, "open", `${host.origin}/taskpane.html`, "Samarbejdsaftale.docx");
  const tp: FrameLocator = page.frameLocator("#tp");
  await expect(tp.getByText("Direkte forbindelse")).toBeVisible();

  /** Prøvekørsel (MIAI_DEMO_PROBE=1): hurtig gennemgang med et stillbillede pr. scene i stedet for pauser. */
  const snap = async (name: string) => {
    if (PROBE) await page.screenshot({ path: test.info().outputPath(`${name}.png`) });
  };

  await pause(page, 3800);
  await stage(page, "card");
  await pause(page, 900);

  // --- 1. Stavning og grammatik som sporede ændringer --------------------------------------
  await stage(page, "caption", "1 · Korrektur", "Ret stavning og grammatik", "Rettelserne lægges ind som sporede ændringer, som du selv godkender.");
  await pause(page, 1400);
  await stage(page, "cursor", 450, 360);
  await pause(page, 700);
  await stage(
    page,
    "select",
    DOC.slice(1, 4).map((p) => p.text),
  );
  await pause(page, 1100);
  await click(page, tp.getByRole("button", { name: "Ret stavning og grammatik", exact: true }));
  await click(page, tp.getByRole("button", { name: "Kør" }));
  const track = tp.getByRole("button", { name: "Indsæt som sporede ændringer" });
  await expect(track).toBeVisible();
  await pause(page, 2600);
  await snap("1-rettelser");
  await stage(page, "track", true);
  await click(page, track);
  await stage(page, "select", []);
  await expect(tp.getByText(/rettelser anvendt som sporede ændringer/)).toBeVisible();
  await stage(page, "cursor", 780, 520);
  await pause(page, 3600);
  await snap("1-sporet");

  // --- 2. Agenten: plan, godkendelse, udførelse --------------------------------------------
  await stage(page, "track", false);
  await stage(page, "caption", "2 · Agent", "Beskriv ændringen med dine egne ord", "MIAI foreslår en plan. Intet ændres, før du har godkendt den.");
  next = SCRIPT.agent;
  await pause(page, 1400);
  await click(page, tp.getByRole("button", { name: "Alle handlinger" }));
  await click(page, tp.getByRole("button", { name: "Agent: redigér dokumentet", exact: true }));
  await type(page, tp.getByRole("textbox", { name: "Hvad skal ændres?" }), "Sæt overskrifter på afsnittene, og saml fristerne i en tabel til sidst");
  await click(page, tp.getByRole("button", { name: "Kør" }));
  const run = tp.getByRole("button", { name: /^Udfør \d+ handling/ });
  await expect(run).toBeVisible();
  await pause(page, 2200);
  // Vis planen: først de første handlinger, så resten og knappen.
  await tp.getByRole("checkbox").nth(2).scrollIntoViewIfNeeded();
  await pause(page, 2200);
  await run.scrollIntoViewIfNeeded();
  await pause(page, 1800);
  await snap("2-plan");
  await click(page, run);
  const done = tp.getByText(/handlinger udført/);
  await expect(done).toBeVisible();
  await done.scrollIntoViewIfNeeded();
  await stage(page, "cursor", 840, 640);
  await pause(page, 4200);
  await snap("2-udfoert");

  // --- 3. Chat med dokumentet ----------------------------------------------------------------
  await stage(page, "caption", "3 · Chat", "Spørg ind til dokumentet", "Svaret kommer fra modellen på din egen maskine. Dokumentet forlader den ikke.");
  next = SCRIPT.chat;
  await pause(page, 1200);
  await click(page, tp.getByRole("tab", { name: "Chat" }));
  const ask = tp.getByPlaceholder("Spørg om dit dokument…");
  await type(page, ask, "Hvilke frister skal kunden overholde?");
  await ask.press("Enter");
  await expect(tp.getByText("fortroligheden også efter aftalens ophør")).toBeVisible();
  await pause(page, 3800);
  await snap("3-chat");

  // --- Afslutning ------------------------------------------------------------------------------
  await stage(
    page,
    "card",
    "Jeres dokumenter. Jeres modeller.",
    "MIAI er åben kildekode under MIT-licensen og virker med Ollama, LM Studio, vLLM og andre OpenAI-kompatible værktøjer.",
    ["Word", "Excel", "PowerPoint", "Outlook"],
    "github.com/mor10w/MIAIPlugin",
  );
  await pause(page, 4500);

  await page.close();
  await host.stop();
  await model.stop();
});
