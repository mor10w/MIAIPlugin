import { expect, test } from "./fixtures/test";

/** Content-Security-Policy (#28): browseren håndhæver politikkens hosts og spærrer andre veje ud. */

/** Prøver et netværkskald og et billede mod en adresse og melder, hvilke CSP-direktiver der blokerede. */
const blocked = (page: import("@playwright/test").Page, url: string) =>
  page.evaluate(async (target) => {
    const hits: string[] = [];
    const onViolation = (e: SecurityPolicyViolationEvent) => hits.push(e.effectiveDirective);
    document.addEventListener("securitypolicyviolation", onViolation);
    await fetch(target).catch(() => {});
    new Image().src = `${target}?billede`;
    await new Promise((r) => setTimeout(r, 300));
    document.removeEventListener("securitypolicyviolation", onViolation);
    // Gælder flere politikker på én gang (header og <meta>), meldes hver overtrædelse én gang pr. politik.
    return [...new Set(hits)].sort();
  }, url);

const csp = (page: import("@playwright/test").Page) =>
  page.evaluate(() => [...document.querySelectorAll<HTMLMetaElement>('meta[http-equiv="Content-Security-Policy"]')].map((m) => m.content));

test("direkte tilstand: kald og billeder til hosts uden for politikken blokeres af browseren; modellen kan stadig kaldes", async ({ open }) => {
  const tp = await open({ host: "Word", selection: "En tekst." });
  // CSP'en er sat ud fra standardpolitikken (kun maskinen selv) og begrænser ikke scripts.
  const [policy] = await csp(tp.page);
  expect(policy).toContain("connect-src 'self' http://127.0.0.1:* https://127.0.0.1:* http://localhost:*");
  expect(policy).toContain("img-src 'self' data: blob:");
  expect(policy).not.toContain("script-src");

  expect(await blocked(tp.page, "https://udenfor.example.com/data")).toEqual(["connect-src", "img-src"]);
  // Modellen på 127.0.0.1 er tilladt af politikken (kun billedet er spærret – billeder må kun komme fra add-in'et selv).
  expect(await blocked(tp.page, `${tp.model.url}/models`)).toEqual(["img-src"]);

  tp.model.reply = () => ({ stream: ["Kortere tekst."] });
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Kortere tekst.")).toBeVisible();
});

test("en ny host i politikken genindlæser task pane, så browseren tillader den; en fjernet host spærres straks", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.section("Privatliv og politik");
  const hosts = tp.page.getByRole("textbox", { name: /Tilladte hosts/ });
  await hosts.fill("127.0.0.1 localhost ::1 llm.firma.local");
  await tp.page.getByRole("button", { name: "Gem" }).click();
  // Efter genindlæsningen står man i Indstillinger igen med en forklaring, og den nye host er med i CSP'en.
  await expect(tp.page.getByText(/MIAI blev genindlæst, så de nye hosts er tilladt/)).toBeVisible();
  expect((await csp(tp.page))[0]).toContain("http://llm.firma.local:*");
  expect(await blocked(tp.page, "http://llm.firma.local:8000/v1/models")).toEqual(["img-src"]);

  // Fjernes den igen, strammes CSP'en med det samme – uden genindlæsning.
  await tp.section("Privatliv og politik");
  await tp.page.getByRole("textbox", { name: /Tilladte hosts/ }).fill("127.0.0.1 localhost ::1");
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.", { exact: true })).toBeVisible();
  expect(await csp(tp.page)).toHaveLength(2);
  expect(await blocked(tp.page, "http://llm.firma.local:8000/v1/models")).toEqual(["connect-src", "img-src"]);
});

test("bridge: task pane må kun tale med bridgen selv (header fra bridgen og samme politik i add-in'et)", async ({ open }) => {
  const tp = await open({ host: "Word", mode: "bridge", selection: "En tekst." });
  const header = await tp.page.evaluate(async () => (await fetch(location.href)).headers.get("content-security-policy"));
  expect(header).toContain("connect-src 'self';");
  expect((await csp(tp.page))[0]).toBe(header);
  // Heller ikke modellen kan kaldes direkte fra task pane – det gør bridgen.
  expect(await blocked(tp.page, `${tp.model.url}/models`)).toEqual(["connect-src", "img-src"]);
  tp.model.reply = () => ({ stream: ["Kortere tekst."] });
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Kortere tekst.")).toBeVisible();
});
