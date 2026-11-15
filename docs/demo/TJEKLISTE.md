# Kan MIAI det, demovideoen viser?

Demovideoen er optaget mod en simuleret Office med indspillede modelsvar. Denne tjekliste gentager de samme fire forløb i **rigtig Word og Excel med din egen model**, så vi ved, om videoen holder. Regn med en halv time.

Filerne ligger i denne mappe:

- [`Samarbejdsaftale.docx`](Samarbejdsaftale.docx): samme tekst som i videoen, med de tre fejl.
- [`Salg.xlsx`](Salg.xlsx): samme tal som i videoen, med to tomme rækker og uden formatering.

## Før du går i gang

- MIAI er installeret og viser **Forbundet**.
- Notér, hvad du tester med; resultatet afhænger af det:

| | |
|---|---|
| Office-version og platform | |
| MIAI-udgave (appens menu, eller Indstillinger) | |
| Inference-værktøj og model (fx Ollama, `qwen3:14b`) | |

Arbejd på kopier af filerne. Gør hvert forløb **tre gange** (luk uden at gemme imellem): modeller svarer forskelligt fra gang til gang, og én heldig kørsel siger ikke nok.

## 1. Korrektur som sporede ændringer (Word)

1. Åbn `Samarbejdsaftale.docx`, og markér de tre afsnit fra *"Denne aftale beskriver…"* til og med *"…rentelovens bestemelser."*
2. Handlinger → **Ret stavning og grammatik** → **Kør**.
3. Vælg **Indsæt som sporede ændringer**.

**Forventet**

- Modellen finder de tre fejl: `d. 1 januar` → `d. 1. januar`, `en løsning der` → `en løsning, der`, `bestemelser` → `bestemmelser`. Finder den flere eller færre, så notér hvilke.
- MIAI melder *"3 af 3 rettelser anvendt som sporede ændringer."*
- I Word står rettelserne som sporede ændringer under **Gennemse**, og de kan accepteres og afvises enkeltvis.
- Resten af teksten og formateringen er uændret.

| Kørsel | Fejl fundet (af 3) | Lagt ind som sporede ændringer? | Bemærkninger |
|---|---|---|---|
| 1 | | | |
| 2 | | | |
| 3 | | | |

## 2. Agenten i Word: overskrifter og tabel

1. Acceptér ændringerne fra forløb 1 (eller åbn en frisk kopi), og fjern markeringen.
2. Handlinger → **Agent: redigér dokumentet**.
3. Skriv: `Sæt overskrifter på afsnittene, og saml fristerne i en tabel til sidst` → **Kør**.
4. Læs planen, og vælg **Udfør … handlinger**.
5. Vælg til sidst **Fortryd agent-session**.

**Forventet**

- MIAI viser en **plan** med konkrete handlinger, ikke en forklaring på, hvordan du selv gør det.
- Planen sætter en overskrift over hvert af de fire afsnit og en tabel sidst i dokumentet med de tre frister (30 dage, 14 dage, 6 måneder). Andre ord og en anden opdeling er i orden.
- Efter udførelsen: overskrifterne har Words typografi *Overskrift 2*, tabellen har en overskriftsrække, og brødteksten er uændret.
- Alle handlinger står som **Udført**; ingen står som *kunne ikke udføres*.
- Efter **Fortryd** ser dokumentet ud som før.

| Kørsel | Plan i første forsøg? | Handlinger udført (x af y) | Resultat som forventet? | Fortryd virkede? | Bemærkninger |
|---|---|---|---|---|---|
| 1 | | | | | |
| 2 | | | | | |
| 3 | | | | | |

## 3. Chat med dokumentet (Word)

1. Fanen **Chat**. Lad *Brug markering/dokument som kontekst* være slået til.
2. Skriv: `Hvilke frister skal kunden overholde?`

**Forventet**

- Svaret nævner godkendelse inden for **30 dage** og betaling inden for **14 dage**, og det opfinder ikke frister, der ikke står i dokumentet.
- Svaret kommer løbende, og dokumentet ændres ikke.

| Kørsel | Begge frister nævnt? | Noget opfundet? | Tid til første ord | Bemærkninger |
|---|---|---|---|---|
| 1 | | | | |
| 2 | | | | |
| 3 | | | | |

## 4. Agenten i Excel: ryd op, formatér, diagram

1. Åbn `Salg.xlsx`.
2. Handlinger → **Agent: redigér dokumentet**.
3. Skriv: `Fjern de tomme rækker, gør tabellen pæn, og vis omsætningen i et diagram` → **Kør**.
4. Læs planen, og vælg **Udfør … handlinger**.
5. Vælg til sidst **Fortryd agent-session**.

**Forventet**

- MIAI viser en **plan**, ikke en vejledning eller en VBA-makro.
- Efter udførelsen:
  - Række 3 og 6 er væk, og de fem regioner står samlet i række 2–6.
  - Området er en Excel-tabel med en tabeltypografi (overskrift og stribede rækker).
  - Der er et diagram med de fem regioner. Det ser anderledes ud end i videoen; det er Excels eget.
  - Talformat (`kr.`) og datalinjer er med i videoen, men en model vælger dem ikke nødvendigvis. Notér, om de kom med.
- Tallene er de samme som før; ingen værdier er flyttet til forkerte celler.
- Efter **Fortryd**: de tomme rækker er tilbage, og tabelformat og diagram er væk.

| Kørsel | Plan i første forsøg? | Handlinger udført (x af y) | Tomme rækker væk? | Tabel? | Diagram? | Fortryd virkede? | Bemærkninger |
|---|---|---|---|---|---|---|---|
| 1 | | | | | | | |
| 2 | | | | | | | |
| 3 | | | | | | | |

## Hvad resultatet betyder

- **Alt som forventet i mindst to af tre kørsler:** videoen er dækkende for den model og den Office-version. Skriv modellen på, når videoen bruges.
- **Planen er god, men handlinger kan ikke udføres, eller resultatet i dokumentet er forkert:** fejlen ligger i MIAI's kode mod rigtig Office. Det skal rettes, før videoen bruges.
- **Modellen forklarer i stedet for at planlægge, eller planen peger på forkerte celler eller afsnit:** modellen er for svag til opgaven. Prøv en større model, og notér, hvilken der skal til.

## Send tilbage

Når noget ikke går som forventet, så gem:

- Et skærmbillede af planen og af dokumentet bagefter.
- Den instruktion, du skrev, og modellens navn.
- Loggen fra MIAI App: `~/Library/Application Support/MIAI/MIAI.log` (Mac) eller `%APPDATA%\MIAI\MIAI.log` (Windows).

Opret et issue med det, eller læg det i issue #1, som samler afprøvningen i rigtig Office.
