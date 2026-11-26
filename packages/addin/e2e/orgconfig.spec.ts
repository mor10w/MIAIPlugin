import { chunks, expect, test } from "./fixtures/test";

/** Central udrulning (#7, ADR-0015): organisationens miai-config.json på add-in'ets adresse. */

test("organisationens politik er låst og vises som organisationens; standardindstillingerne peger på den interne server", async ({ open }) => {
  const tp = await open({
    host: "Word",
    orgConfig: {
      policy: { allowedHosts: ["127.0.0.1", "llm.firma.local"], webSearch: false },
      settings: { uiLanguage: "da", temperature: 0.2 },
    },
  });
  await tp.tab("Indstillinger");
  await tp.section("Privatliv og politik");
  await expect(tp.page.getByText("Politik: Organisationen (central udrulning)")).toBeVisible();
  // Låst: tilladte hosts kan ikke redigeres, og de er organisationens.
  const hosts = tp.page.getByRole("textbox", { name: /Tilladte hosts/ });
  await expect(hosts).toBeDisabled();
  await expect(hosts).toHaveValue(/llm\.firma\.local/);

  // En endpoint uden for organisationens hosts afvises ved gem.
  await tp.section("Forbindelse");
  await tp.page.getByLabel("Basis-URL (inkl. /v1)").fill("https://ekstern.example.com/v1");
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText(/ikke tilladt af politikken/)).toBeVisible();
});

test("organisationens faste instruktioner (#78) vises skrivebeskyttet og sendes med før brugerens", async ({ open }) => {
  const tp = await open({
    host: "Word",
    selection: "En lang og indviklet tekst.",
    orgConfig: { policy: { allowedHosts: ["127.0.0.1"], instructions: { word: "Skriv altid i du-form.", excel: "Kun til Excel." } } },
  });
  tp.model.reply = () => ({ stream: chunks("Kort tekst.") });
  await tp.tab("Indstillinger");
  await tp.section("Faste instruktioner");

  // Organisationens instruktion kan læses, men ikke ændres – heller ikke efter gem.
  const org = tp.page.getByRole("textbox", { name: "Fra din organisation i Word" });
  await org.scrollIntoViewIfNeeded();
  await expect(org).toHaveValue("Skriv altid i du-form.");
  await expect(org).not.toBeEditable();
  await expect(tp.page.getByText(/Fastsat af din organisation og kan ikke ændres/).first()).toBeVisible();
  // Programmer uden instruktion fra organisationen har kun brugerens felt.
  await expect(tp.page.getByRole("textbox", { name: "Fra din organisation i Excel" })).toHaveValue("Kun til Excel.");
  await expect(tp.page.getByRole("textbox", { name: "Fra din organisation i PowerPoint" })).toHaveCount(0);

  await tp.page.getByRole("textbox", { name: "Instruktioner i Word" }).fill("Brug korte sætninger.");
  await tp.page.getByRole("button", { name: "Gem" }).click();
  await expect(tp.page.getByText("Gemt.")).toBeVisible();
  await expect(org).toHaveValue("Skriv altid i du-form.");

  await tp.tab("Handlinger");
  await tp.run("Forenkl");
  await expect(tp.page.getByText("Kort tekst.")).toBeVisible();
  const system = tp.model.lastRequest!.messages[0].content;
  expect(tp.model.lastRequest!.messages[0].role).toBe("system");
  expect(system).toContain("Skriv altid i du-form.");
  expect(system.indexOf("Brug korte sætninger.")).toBeGreaterThan(system.indexOf("Skriv altid i du-form."));
  expect(system).not.toContain("Kun til Excel.");
});

test("en ulæselig miai-config.json låser politikken (fail closed) i stedet for at åbne for mere", async ({ open }) => {
  const tp = await open({ host: "Word", orgConfig: "{ dette er ikke json" });
  await tp.tab("Indstillinger");
  await tp.section("Privatliv og politik");
  await expect(tp.page.getByText("Politik: Organisationen (central udrulning)")).toBeVisible();
});

test("uden miai-config.json virker add-in'et som før (brugerens egen politik)", async ({ open }) => {
  const tp = await open({ host: "Word" });
  await tp.tab("Indstillinger");
  await tp.section("Privatliv og politik");
  await expect(tp.page.getByText("Politik: standard (kun denne maskine)")).toBeVisible();
});
