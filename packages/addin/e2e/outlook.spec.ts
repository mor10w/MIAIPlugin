import { expect, test } from "./fixtures/test";

/** Outlook (#7, ADR-0015): mail i læsetilstand og udkast i skrivetilstand. */

const MAIL = {
  mode: "read" as const,
  subject: "Møde i næste uge",
  from: { name: "Bo Jensen", email: "bo@firma.dk" },
  body: "Hej\n\nKan vi mødes onsdag kl. 10 om budgettet? Jeg har brug for et svar inden fredag.\n\nVh Bo",
};

test("læsetilstand: kun opgaver, der giver mening for en mail; svaret åbnes som et udkast med formatering", async ({ open }) => {
  const tp = await open({ host: "Outlook", mail: MAIL });
  await expect(tp.page.getByRole("button", { name: "Svar på mail", exact: true })).toBeVisible();
  await expect(tp.page.getByRole("button", { name: "Opsummér", exact: true })).toBeVisible();
  // Ingen Excel-, PowerPoint- eller agentopgaver.
  for (const name of ["Lav formel ud fra beskrivelse", "Tekst til slides", "Agent: redigér dokumentet", "Design: master og layout"])
    await expect(tp.page.getByRole("button", { name, exact: true })).toHaveCount(0);

  tp.model.reply = () => ({ stream: ["Hej Bo\n\nTak – **torsdag kl. 10** passer bedre, hvis det er i orden.\n\nVh"] });
  await tp.page.getByRole("button", { name: "Svar på mail", exact: true }).click();
  await tp.page.getByRole("textbox", { name: /Hvad skal svaret sige/ }).fill("Foreslå torsdag i stedet");
  await tp.page.getByRole("button", { name: "Kør" }).click();
  await expect(tp.page.getByRole("button", { name: "Svar alle" })).toBeVisible();

  // Modellen fik mailen med emne og afsender – som data.
  const user = tp.model.lastRequest!.messages.at(-1)!.content;
  expect(user).toContain("What the reply should say: Foreslå torsdag i stedet");
  expect(user).toContain("<document>\nSubject: Møde i næste uge\nFrom: Bo Jensen <bo@firma.dk>\n\nHej");
  expect(tp.model.lastRequest!.messages[0].content).toContain("Microsoft Outlook");

  await tp.page.getByRole("button", { name: "Svar", exact: true }).click();
  await expect(tp.page.getByText("Svaret er åbnet som et udkast.")).toBeVisible();
  const reply = await tp.find<{ all: boolean; html: string }>("replyForm");
  expect(reply).toMatchObject({ all: false });
  expect(reply!.html).toContain("<strong>torsdag kl. 10</strong>");
});

test("skrivetilstand: omskriv markeringen og indsæt den som HTML i udkastet", async ({ open }) => {
  const tp = await open({
    host: "Outlook",
    mail: { mode: "compose", subject: "Tilbud", body: "Hej Anna\n\nhermed tilbuddet, det er billigt\n\nVh", selection: "hermed tilbuddet, det er billigt" },
  });
  tp.model.reply = () => ({ stream: ["Vedlagt finder du vores tilbud – en **konkurrencedygtig** pris."] });
  await tp.run("Omskriv i anden tone");
  await expect(tp.page.getByText(/konkurrencedygtig/)).toBeVisible();
  expect(tp.model.lastRequest!.messages.at(-1)!.content).toContain("hermed tilbuddet, det er billigt");
  await tp.page.getByRole("button", { name: "Erstat" }).click();
  const set = await tp.find<{ text: string; coercionType: string }>("setSelectedData");
  expect(set).toMatchObject({ coercionType: "html" });
  expect(set!.text).toContain("Vedlagt finder du vores tilbud");
});

test("chat om mailen: forslag til mails, og hele mailen er konteksten", async ({ open }) => {
  const tp = await open({ host: "Outlook", mail: MAIL });
  await tp.tab("Chat");
  await expect(tp.page.getByRole("button", { name: "Er der frister eller åbne spørgsmål?" })).toBeVisible();
  tp.model.reply = () => ({ stream: ["Ja: svar inden fredag."] });
  await tp.page.getByRole("button", { name: "Er der frister eller åbne spørgsmål?" }).click();
  await expect(tp.page.getByText("Ja: svar inden fredag.")).toBeVisible();
  expect(tp.model.lastRequest!.messages[0].content).toContain("Jeg har brug for et svar inden fredag.");
});
