# ADR-0008: Distribution af solgte løsninger – hostet via M365-admin eller lokal app

- Status: Vedtaget
- Dato: 2026-11-02
- Relateret: [ADR-0007](0007-bridge-valgfri-direkte-forbindelse.md)

## Kontekst
For en kunde, der har købt løsningen, er "driv selv en HTML-side" (GitHub Pages) og manuel sideloading uden for comfort zone. Microsofts dokumentation ("Deploy and publish Office Add-ins", opdateret 2025-10) fastslår, at sideloading kun er til udvikling/afprøvning. Production-vejene for web-add-ins er:

1. **Microsoft Marketplace** (Office Store) – offentlig distribution; leverandør hoster web-indholdet.
2. **Integrated apps portal i M365-admin center** – kunden (ISV-køb) eller organisationen deployer add-in'et til brugere/grupper. "No client configuration required" – endbruger ser bare add-in'et i Office.
3. **SharePoint-katalog / Centralized Deployment** – on-prem-organisationer og sovereign clouds.

I alle tre tilfælde hostes add-in'ets web-indhold af **leverandøren** (egent domæne, fx Azure App Service/Static Web Apps bag CDN) – kunden driver intet. COM/VSTO-add-ins er ifølge Microsoft i maintenance mode ("no new development") og Windows-only, så de afvises som fundament for nye funktioner.

## Beslutning
Vi sælger to udgaver af samme kodebase (selv ADR-0007's to tilstande er allerede implementeret):

| | **Udgave A – "hostet"** (M365-kunder) | **Udgave B – "lokal app"** (fx .dmg/.msi) |
|---|---|---|
| Add-in'ets filer | Hostes af os på eget domæne (`scripts/make-manifest.mjs hosted https://…`) | Serveres af MIAI Bridge på `https://localhost:44321` |
| Installation hos kunden | IT-adminen uploader manifestet i Integrated apps portal (én gang) – evt. som SaaS-offer i Commercial Marketplace | Kunden installerer MIAI App (Electron-menu-bar-app, `packages/app`), som kører bridgen i en utility-process |
| Tilstand | Direkte (webview → model) | Bridge (samme origin – ingen CORS) |
| Forudsætninger hos kunden | M365-admin + `OLLAMA_ORIGINS` for vores domæne (evt. via GPO/Intune) | Ingen – alt kører lokalt, ingen vendor-cloud, ingen internet |
| Politik/audit | Brugerens (loopback-standard) | Maskinpolitik + audit-log kan leveres som standard |
| Opdateringer | Vi deployer; adminen godkender manifest-ændringer, der rører rettigheder | App-opdatering (senere: auto-update via zip-mål) |

**Udgave B er standard for salgs løsninger**, hvor fortrolighed ("dokumenter forlader ikke maskinen") er kernekravet, og den anbefales også til kunder uden M365-admin. Udgave A er vejen, når kunden har en M365-tenant og ønsker central udrulning.

Teknisk indhold i `packages/app` (Udgave B):
- Electron-menu-bar-app, der kører `@miai/bridge` i en Electron utility-process (ingen Node-installation hos brugeren). Bridgen kan ikke køre i main-processen, fordi TLS-håndtryk hænger der.
- Genererer eget TLS-certifikat for `localhost` (`~/Library/Application Support/MIAI/tls`). Gyldigheden er **825 dage**, som er macOS' maksimum for TLS-servercertifikater; certifikatet fornyes automatisk 30 dage før udløb. Appen tilbyder at betro det i login-keychain (`security add-trusted-cert`).
- Opsætningsdialogen vises kun, når noget mangler: certifikatet er ikke betroet (tjekkes med `security verify-cert`), eller manifestet ligger ikke i Office.
- Sideloader bridge-manifestet i Word/Excel/PowerPoint automatisk (`~/Library/Containers/com.microsoft.*/Data/Documents/wef`).
- Indstilling af model-URL og valg af model sker **i add-in'ets fanen Indstillinger** (endpointer + standardmodel), som i øvrigt gælder begge udgaver.

## Konsekvenser
+ Kunden driver aldrig web-indhold; installation er enten én admin-upload eller ét setup.
+ Udgave B har ingen afhængighed af vores infrastruktur – løsningen kører uforandret, hvis vi lukker servere.
+ To distributionsspor at vedligeholde (hosting + app), men samme motor og prompts.
- Udgave A: webviewet kalder localhost fra et offentligt origin – "Local Network Access"-tilladelser (Chromium) skal afprøves pr. Office-version; CORS kræver `OLLAMA_ORIGINS` for vores domæne på kundens maskiner.
- Udgave B: Electron er tungt (~150 MB) og kører i menuen. Certifikatet skal betroes én gang og igen ved fornyelse (ca. hvert andet år).
- Uden kode-signering advarer Gatekeeper/SmartScreen ved første start – signering (Apple Developer / code-signing-certifikat) er en forudsætning for seriøs salgsudlevering.
