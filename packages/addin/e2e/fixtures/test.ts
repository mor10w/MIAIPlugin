import { test as base, expect, type Page } from "@playwright/test";
import { fakeOfficeScript, type FakeOfficeOptions, type OfficeCall } from "./office";
import { MockModel, startBridge, startStaticHost, type MockNative } from "./servers";

export { expect };

export interface TaskpaneOptions extends FakeOfficeOptions {
  /** direct = add-in hostet statisk, kalder modellen selv; bridge = serveret af MIAI Bridge. */
  mode?: "direct" | "bridge";
  /** Mock-modellen sender CORS-headere (standard: ja). */
  cors?: boolean;
  /** Ekstra felter i endpoint-konfigurationen, fx { presencePenalty: 1.5 }. */
  endpoint?: Record<string, unknown>;
  /** Ekstra indstillinger, fx { webSearch: true }. */
  settings?: Record<string, unknown>;
  /** Modeller, mock-modellen udbyder på /v1/models (standard: ["qwen3:4b"]). */
  models?: string[];
  /** Fast port til mock-modellen (fx 11434 for at ligne Ollama). Standard: tilfældig. */
  modelPort?: number;
  /** Mock-modellen svarer også som Ollamas eller LM Studios eget API (#73). */
  native?: MockNative;
  /** Direkte tilstand: organisationens miai-config.json på add-in'ets adresse (#7). Streng = råt indhold (fx ugyldig JSON). */
  orgConfig?: Record<string, unknown> | string;
  /** Bridge: åbn task pane uden adgangstoken (som med et gammelt manifest). */
  omitToken?: boolean;
}

export interface Taskpane {
  page: Page;
  model: MockModel;
  /** Bridge-tilstand: filer, bridgen er blevet bedt om at åbne (F15, "Åbn fil"). */
  opened: string[];
  /** Det, adapterne har gjort ved det simulerede dokument. */
  calls(): Promise<OfficeCall[]>;
  /** Første registrerede kald af en bestemt slags (fx "picture"), typet efter behov. */
  find<T = Record<string, unknown>>(kind: string): Promise<T | undefined>;
  /** Vælger en fane i task pane. */
  tab(name: "Handlinger" | "Egne opgaver" | "Chat" | "Indstillinger"): Promise<void>;
  /** Åbner en sektion under Indstillinger (accordion), hvis den er lukket. */
  section(name: "Forbindelse" | "Modeller" | "Svar" | "Privatliv og politik" | "Viden fra mine mapper" | "Excel-funktioner" | "Generelt"): Promise<void>;
  /** Vælger en handling og trykker Kør. */
  run(action: string): Promise<void>;
}

type Open = (opts: TaskpaneOptions) => Promise<Taskpane>;

/**
 * `open` starter mock-model og hosting (direkte eller rigtig bridge), indlæser task pane med
 * simuleret Office.js og rydder op efter testen. Alle porte er dynamiske, så en lokal Ollama
 * på 11434 ikke forstyrrer.
 */
export const test = base.extend<{ open: Open }>({
  open: async ({ page }, use) => {
    const cleanup: (() => Promise<void>)[] = [];
    const errors: string[] = [];
    page.on("pageerror", (e) => errors.push(e.message));

    await use(async (opts) => {
      const model = await new MockModel({ cors: opts.cors, models: opts.models, port: opts.modelPort, native: opts.native }).start();
      cleanup.push(() => model.stop());

      await page.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: fakeOfficeScript(opts) }));

      let origin: string;
      let query = "";
      let opened: string[] = [];
      if (opts.mode === "bridge") {
        const endpointExtra = opts.endpoint ?? {};
        const bridge = await startBridge(model.url, {
          ...opts.settings,
          ...(Object.keys(endpointExtra).length ? { endpoints: [{ id: "mock", name: "Mock", baseUrl: model.url, ...endpointExtra }] } : {}),
        });
        cleanup.push(bridge.stop);
        origin = bridge.origin;
        opened = bridge.opened;
        if (!opts.omitToken) query = `?token=${encodeURIComponent(bridge.token)}`;
      } else {
        const host = await startStaticHost();
        cleanup.push(host.stop);
        origin = host.origin;
        // Direkte tilstand: indstillingerne ligger i add-in'ets localStorage – peg dem på mock-modellen.
        const settings = { endpoints: [{ id: "mock", name: "Mock", baseUrl: model.url, ...opts.endpoint }], activeEndpoint: "mock", ...opts.settings };
        await page.addInitScript((s) => localStorage.setItem("miai.settings.v1", s), JSON.stringify(settings));
      }

      if (opts.orgConfig !== undefined) {
        const body = typeof opts.orgConfig === "string" ? opts.orgConfig : JSON.stringify(opts.orgConfig);
        await page.route(`${origin}/miai-config.json`, (r) => r.fulfill({ contentType: "application/json", body }));
      }
      // Et manifest med regnearksfunktionerne åbner task pane med ?functions=1 (#80).
      if (opts.customFunctions) query += `${query ? "&" : "?"}functions=1`;
      await page.goto(`${origin}/taskpane.html${query}`);
      // I en skjult funktions-runtime starter panelet aldrig, så der er intet at vente på.
      if (!opts.hiddenRuntime) await expect(page.getByText(opts.mode === "bridge" ? "Via MIAI Bridge" : "Direkte forbindelse")).toBeVisible();

      return {
        page,
        model,
        opened,
        calls: () => page.evaluate(() => (window as unknown as { __miai: { calls: OfficeCall[] } }).__miai.calls),
        find: async <T>(kind: string) =>
          (await page.evaluate(() => (window as unknown as { __miai: { calls: OfficeCall[] } }).__miai.calls)).find((c) => c.kind === kind) as T | undefined,
        tab: async (name) => page.getByRole("tab", { name }).click(),
        section: async (name) => {
          const header = page.getByRole("button", { name, exact: true });
          if ((await header.getAttribute("aria-expanded")) !== "true") await header.click();
        },
        run: async (action) => {
          // En åben handling vises i sit eget trin – gå tilbage til listen først.
          const back = page.getByRole("button", { name: "Alle handlinger" });
          if (await back.isVisible()) await back.click();
          await page.getByRole("button", { name: action, exact: true }).click();
          await page.getByRole("button", { name: "Kør" }).click();
        },
      };
    });

    // Luk siden først, så browserens forbindelser er væk, før serverne stoppes.
    await page.close();
    for (const c of cleanup.reverse()) await c();
    // Uventede JavaScript-fejl i task pane får testen til at fejle, også når selve flowet lykkedes.
    expect(errors, "fejl i task pane").toEqual([]);
  },
});

/** Hjælper til at bygge et streamet svar af en tekst (små bidder, som en rigtig model). */
export function chunks(text: string, size = 6): string[] {
  return text.match(new RegExp(`[\\s\\S]{1,${size}}`, "g")) ?? [];
}
