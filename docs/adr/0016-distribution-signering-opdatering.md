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
  1. Ret `version` i `packages/app/package.json`.
  2. Tag `vX.Y.Z`.
  3. Release-workflowet bygger, signerer og laver en GitHub Release.
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
