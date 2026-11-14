# ADR-0016: Distribution, kodesignering og opdatering af MIAI App

- **Status:** Foreslået. Bedste bud på signering; afventer bekræftelse og køb af certifikater.
- **Dato:** 2026-11-04
- **Issue:** #5

## Kontekst

#5 kræver:
- en installer klar til betatest
- en beslutning om kodesignering (certifikat og udgiver), jf. PLAN.md §9, spørgsmål 3
- en beskrevet opdateringsstrategi

Situationen før:
- MIAI App fandtes kun til Mac: usigneret, uden automatisk opdatering og med `/usr/bin/openssl` og macOS-stier hårdkodet.
- Den hostede udgave (GitHub Pages, M365 Admin Center og nu Outlook, ADR-0015) kræver ingen installer.

## Beslutning

### 1. Distributionsveje

| Udgave | Til | Installation |
|---|---|---|
| Hostet add-in | Organisationer med M365 | M365 Admin Center → Integrated apps (ADR-0015). Ingen installer. |
| MIAI App til Mac | SMB og private | `.dmg` (Apple Silicon og Intel) |
| MIAI App til Windows | SMB og private | `MIAI-Setup-X.Y.Z-x64.exe` (NSIS, pr. bruger, ingen admin). Senere via winget. |

### 2. MIAI App til Windows

Samme Electron-app og bridge som på Mac. Det platformsspecifikke ligger i `packages/app/platform.js`:

| | macOS | Windows |
|---|---|---|
| Office (Word/Excel/PowerPoint) | Manifestet i hver apps `wef`-mappe | Stien i `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` (pr. bruger, ingen admin) |
| Certifikat | Login-nøgleringen (`security`, ingen dialog) | Brugerens rodlager (`certutil -user`, Windows viser en dialog, brugeren godkender) |
| Certifikatet laves med | `selfsigned` (ren JavaScript) på begge platforme, 825 dage. Fornyes 30 dage før udløb, og brugeren bedes betro det igen. | |
| Start ved login | Login-emner (SMAppService) | `setLoginItemSettings` (registreringsdatabasens Run-nøgle) |
| Proceslinje-ikon | Skabelonikon | Farvet ikon |

Outlook installeres fra fil på begge platforme, for Outlook har ingen sideload-mappe.

### 3. Kodesignering (bedste bud)

- **Udgiver:** brug **virksomhedens registrerede navn (CVR)** frem for et personnavn. Navnet vises i Gatekeeper, SmartScreen og M365 Admin Center, og en virksomhed giver adgang til de billigste signeringsløsninger nedenfor. Indtil da står "M. Wiingaard" i manifesterne.
- **macOS:** Apple Developer Program (ca. 99 USD/år).
  - Som organisation kræver det et D-U-N-S-nummer (gratis); som person kan det oprettes i eget navn.
  - Certifikat: **Developer ID Application**.
  - Notarisering med en App Store Connect API-nøgle.
  - Det er nødvendigt både for at Gatekeeper ikke advarer, og for at automatisk opdatering virker på Mac (Squirrel.Mac kræver en signeret app).
- **Windows:** **Azure Trusted Signing** (ca. 10 USD/md., Basic).
  - Det er billigst og kræver ingen hardware-token.
  - Det er tilgængeligt for organisationer i bl.a. EU efter en identitetsvalidering.
  - EV-certifikater giver ikke længere straks SmartScreen-omdømme, så de er ikke merprisen værd.
  - **Plan B:** et OV-certifikat fra en CA med nøgle i cloud-HSM (ca. 200–400 USD/år).
- Priser og betingelser er fra min viden og bør bekræftes ved køb.
- **Pipelinen er klar:** `.github/workflows/release.yml` signerer og notariserer, så snart secrets er lagt ind (navnene står i workflowets top). Indtil da bygges der usigneret med en tydelig advarsel i kørslen.

### 4. Opdatering

- **electron-updater** henter fra et "generic" feed, `https://mor10w.github.io/MIAIPlugin/updates/` (`latest.yml` og `latest-mac.yml`).
  - Den tjekker ved start og hver 6. time og henter i baggrunden.
  - Opdateringen installeres ved næste genstart, eller med det samme fra menuen ("Genstart og opdater til X").
- **Udgivelse:**
  1. Push til `main` (app-kode): patch-versionen bumpes automatisk, og der tages et tag.
  2. Eller: ret `version` i `packages/app/package.json`, hvis en bestemt version skal udgives.
  3. Release-workflowet bygger, signerer og laver en GitHub Release (tagget trigget den).
  4. Pages kopierer installationsfiler og feed fra seneste release ud, for repoet er privat, og kunderne kan ikke hente fra en privat release.
- **Kanaler:** én (`latest`). En betakanal kan tilføjes med electron-updaters `channel` og pre-release-tags, når der er behov.
- **winget:** skabelon i `packaging/winget/`. Den indsendes, når installeren er signeret, fordi winget-pkgs i praksis kræver en signeret, offentligt tilgængelig installer.

### 5. Gentagen betroelse af certifikatet

- Hvert 825. dag (macOS' grænse; Windows følger samme) laves et nyt certifikat.
- Ved start opdager appen, at det ikke er betroet, og spørger én gang.
- På Mac sker det uden dialog. På Windows viser systemet sin sikkerhedsdialog, og velkomstdialogen siger "vælg Ja".

## Konsekvenser

- **Indtil der er certifikater:**
  - På Mac skal brugeren højreklikke → Åbn første gang.
  - På Windows skal brugeren vælge "Flere oplysninger → Kør alligevel".
  - Automatisk opdatering virker på Windows, men ikke på Mac.
  - Fint til betatest; ikke til salg.
- **Pages som filserver:** GitHub Pages har en grænse på ca. 1 GB. En release fylder ca. 0,5 GB: dmg og zip til to Mac-arkitekturer samt én Windows-installer. Ved salg bør filerne ligge på et CDN eller organisationens egen server. Feed-URL'en ændres i `build.publish`.
- **Windows-installeren er testet i CI's pakning, ikke på en rigtig pc.** Office' registreringsnøgle, `certutil`-dialogen og SmartScreen skal afprøves; tjeklisten står i UDVIKLING.md.
- **Sideloading** (wef-mappe/registreringsdatabase) er ifølge Microsoft til udvikling. For organisationer er central udrulning (ADR-0015) den anbefalede vej; MIAI App er til dem uden M365-administration.

## Tillæg: udviklerudgave ved hvert push til `main`

En rigtig udgivelse kræver et versionsnummer og et tag. Under udviklingen er der brug for at kunne afprøve det, der lige er merget, uden den ceremoni.

- **Hvad:** `.github/workflows/dev-build.yml` bygger en `.dmg` til Apple Silicon ved hvert push til `main`, der rører appen (ikke ved rene dokumentations-, eval- eller e2e-ændringer).
- **Version:** `<version i package.json>-dev.<kørselsnummer>`, fx `0.1.0-dev.42`. Intet git-tag og ingen ændring i repoet; nummeret sættes kun i bygget.
- **Hvor:** én rullende forhåndsudgivelse med tagget `dev`, som erstattes hver gang. Pages kopierer den til `<pages-url>/dev/` sammen med `dev.json` (version, filnavn, commit, dato), fordi repoet er privat.
- **I appen:** menupunktet *Hent udviklerudgave X…* vises, når `dev.json` peger på et nyere byg end det installerede. Det åbner hentningen i browseren; appen installerer ikke selv. Kun en udviklerversion og et rent `.dmg`-filnavn i feedets egen mappe accepteres, så feedet ikke kan sende brugeren andre steder hen.
- **To spor, der ikke blandes:** en installeret udviklerudgave følger kun udviklersporet (electron-updater er slået fra i den), og "seneste release" på Pages omfatter ikke forhåndsudgivelser. `MIAI_NO_UPDATES=1` slår begge fra.
- **Signering:** ad hoc (`identity: "-"`) uden hardened runtime. Apple Silicon kræver en gyldig signatur for overhovedet at starte en app, og hardened runtime uden et Team ID afviser Electrons egne biblioteker. Workflowet kontrollerer signaturen og starter appen, før noget udgives.
- **Kun Mac på Apple Silicon indtil videre.** Intel-Mac og Windows er udeladt for at holde byggetiden nede; Windows er udskudt til #40.

**Konsekvenser:**

- Gatekeeper kender ikke en ad hoc-signatur. Efter hver hentning skal appen godkendes under Systemindstillinger → Anonymitet & sikkerhed → *Åbn alligevel*.
- En ad hoc-signatur er ny for hvert byg. macOS kan derfor spørge igen om adgang til nøgleringen og om tilladelser, der er givet til den forrige udgave. Det forsvinder med et Developer ID-certifikat (afsnit 3).
- Mac-runnere koster ti gange så mange Actions-minutter som Linux i et privat repo. Derfor stifiltrene, og derfor annulleres et ældre byg, når et nyt push kommer.
- Udviklerudgaven fylder ca. 0,1 GB af Pages' 1 GB oven i en release.
- Udviklerudgaven er til afprøvning. Den er ikke testet ud over CI og skal ikke gives til kunder.

## Tillæg 2: Automatisk release ved push til `main`

Hver gang koden ændres, skal kunderne ikke vente på, at nogen husker at bump version og tage et tag.

- **Hvad:** `.github/workflows/auto-release.yml` udgiver en release ved hvert push til `main`, der rører app-koden (samme stifiltre som udviklerudgaven).
- **Version:** bumpes patchen automatisk i `packages/app/package.json` (og lockfilen), committes til `main` af github-actions[bot] og tagges. Ændrer push'et versionen selv, udgives den præcist – minor/major styres stadig manuelt. En version, der ikke er ren semver (X.Y.Z), udgives aldrig automatisk.
- **Uden løkke og uden race:** en commit, der allerede er tagget ved tippet på `main`, udløser intet nyt release; kørsler er sekventielle (concurrency), og udgivelsen sker altid fra seneste tip, så parallelle push'er lander i én release.
- **Konsekvens:** versionsnumre og GitHub Releases flyver under aktiv udvikling (hver release fylder ca. 0,5 GB i repoet; Pages serverer kun den seneste). Når produktet modner, kan man skifte til manuelt eller kræve et review på releases.
