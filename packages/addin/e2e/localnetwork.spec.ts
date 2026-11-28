import type { Page } from "@playwright/test";
import { expect, test } from "./fixtures/test";

/**
 * Lokal netværksadgang (#74): Chrome 142+ og Edge 143+ kræver brugerens tilladelse, før et hostet
 * add-in må kalde modellen på maskinen. Testens Chromium håndhæver ikke kravet (siden ligger selv
 * på localhost), så en blokering simuleres: alle kald til http://127.0.0.1 afbrydes, som når browseren
 * blokerer dem, og Permissions API'et overskrives med det svar, browseren ville give.
 */

const TITLE = "Browseren blokerer adgangen til din lokale model.";
const HELP = /Browseren har afvist, at add-in'et må kalde adresser på denne maskine.*Webstedsindstillinger.*Tillad/;

async function blockLoopback(page: Page) {
  await page.route(/^http:\/\/127\.0\.0\.1[:/]/, (r) => r.abort("failed"));
}

test("afvist tilladelse: den specifikke besked vises i banner, status, ved kørsel og i guiden", async ({ open, page }) => {
  await blockLoopback(page);
  await page.addInitScript(() => {
    const original = navigator.permissions.query.bind(navigator.permissions);
    navigator.permissions.query = (d: PermissionDescriptor) =>
      (d.name as string) === "loopback-network" ? Promise.resolve({ state: "denied" } as PermissionStatus) : original(d);
  });
  const tp = await open({ host: "Word", selection: "Tekst" });

  // Banneret forklarer blokeringen; den generelle besked og "Kom i gang" vises ikke.
  const banner = tp.page.locator(".fui-MessageBar").filter({ hasText: TITLE });
  await expect(banner).toContainText(HELP);
  await expect(tp.page.getByRole("button", { name: "Status: Blokeret af browseren" })).toBeVisible();
  await expect(tp.page.getByText("Ingen forbindelse til en model endnu.")).toHaveCount(0);
  await expect(tp.page.getByRole("button", { name: "Kom i gang" })).toHaveCount(0);

  // En opgave fejler med samme overskrift og adressen – ikke "Kunne ikke forbinde" med CORS-vejledning.
  await tp.run("Forenkl");
  // (Afgrænset til beskederne: statussens tooltip viser samme tekst, når musen står over den.)
  const notices = tp.page.locator(".fui-MessageBar").filter({ hasText: TITLE });
  await expect(notices).toHaveCount(2);
  await expect(notices.last()).toContainText(/Browseren blokerer add-in'ets adgang til http:\/\/127\.0\.0\.1:\d+\/v1 \(lokal netværksadgang\)/);
  await expect(tp.page.getByText(/Kunne ikke forbinde|tillader CORS/)).toHaveCount(0);

  // Guiden viser blokeringen én gang i stedet for "Intet værktøj fundet".
  await tp.tab("Indstillinger");
  const guide = tp.page.locator(".fui-Card").filter({ hasText: "Find lokale værktøjer" });
  await expect(guide.getByText("Blokeret af browseren")).toBeVisible();
  await expect(guide).toContainText(HELP);
  await expect(guide.getByText(/Intet værktøj fundet/)).toHaveCount(0);
  await expect(guide.getByText(/Ikke fundet/)).toHaveCount(0);
  expect(tp.model.requests).toHaveLength(0);
});

test("uden Permissions API (ældre browsere, Office's webviews): den generelle besked som før", async ({ open, page }) => {
  await blockLoopback(page);
  await page.addInitScript(() => Object.defineProperty(navigator, "permissions", { value: undefined }));
  const tp = await open({ host: "Word", selection: "Tekst" });

  await expect(tp.page.getByText("Ingen forbindelse til en model endnu.")).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Status: Ingen forbindelse til modellen" })).toBeVisible();
  await tp.run("Forenkl");
  await expect(tp.page.locator(".fui-MessageBarBody").filter({ hasText: "Kunne ikke forbinde" })).toContainText(/tillader CORS fra https:\/\/localhost:\d+/);

  await tp.page.getByRole("button", { name: "Kom i gang" }).click();
  const guide = tp.page.locator(".fui-Card").filter({ hasText: "Find lokale værktøjer" });
  await expect(guide.getByText(/Intet værktøj fundet/)).toBeVisible();
  await expect(tp.page.getByText(TITLE)).toHaveCount(0);
  await expect(tp.page.getByText("Blokeret af browseren")).toHaveCount(0);
});
