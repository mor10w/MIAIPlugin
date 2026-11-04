import { expect, test } from "./fixtures/test";

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
