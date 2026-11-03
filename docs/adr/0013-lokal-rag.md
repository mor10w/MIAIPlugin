# ADR-0013: Lokal RAG over egne mapper – node:sqlite og vektorsøgning i bridgen

- **Status:** Besluttet
- **Dato:** 2026-11-03
- **Issue:** #8

## Kontekst

F15: chatten skal kunne svare ud fra brugerens egne filer (politikker, referater, rapporter) og vise, hvor svaret kommer fra. Krav og forbehold fra PLAN.md og #8:

- Embeddings fra samme OpenAI-kompatible endpoint (`/v1/embeddings`). Ikke alle værktøjer understøtter det.
- Kun bridgen kan læse brugerens filer. Browseren i direkte tilstand kan ikke.
- Formaterne docx, pdf, xlsx, pptx og md kræver parsere.
- En organisation skal kunne bestemme, hvilke mapper der må indekseres, og audit-loggen skal registrere indeksering, men kun metadata.
- Kildehenvisninger: Office-webviewet kan ikke åbne `file://`-stier.

PLAN.md foreslog SQLite + sqlite-vec.

## Beslutning

1. **Lagring i Nodes indbyggede SQLite (`node:sqlite`)**, i én fil `knowledge.db` ved siden af `config.json`. Der er ingen native moduler at bygge, signere eller pakke til Windows, Mac og Electron. Kræver Node 22.13 eller nyere. Mangler `node:sqlite`, slås kun vidensindekset fra, med en forklaring; resten af bridgen virker.
2. **Ingen sqlite-vec – vektorsøgning med brute force i hukommelsen.**
   - Embeddings gemmes normaliserede som Float32-blobs og læses ind ved første søgning.
   - Cosinus-lighed er så et prikprodukt. 100.000 uddrag × 1.024 dimensioner tager få millisekunder, hvilket rigeligt dækker personlige og afdelingsmapper (loft: 10.000 filer).
   - Gevinsten er, at der ikke skal indlæses en native SQLite-udvidelse, som skulle signeres og notariseres på macOS.
   - Bliver indekserne meget større (fx fælles drev), kan sqlite-vec eller en ANN-indeksering tilføjes bag `KnowledgeStore.search` uden at ændre resten.
3. **Udtræk i ren JavaScript**:
   - Word: `mammoth`.
   - PDF: `unpdf` (pdf.js), med sidetal i teksten.
   - Excel og PowerPoint: læses direkte fra Office Open XML med `jszip`, ark for ark og slide for slide.
   - Markdown, tekst og CSV læses som de er.
   - Ulæselige filer registreres pr. fil og stopper ikke indekseringen.
4. **Opdeling og embeddings**:
   - Uddrag på ca. 1.200 tegn ved afsnits- og sætningsgrænser, med 200 tegns overlap.
   - Filnavnet står foran hvert uddrag, der sendes til embedding.
   - Kendte modeller får deres præfikser for søgning og dokument (nomic, e5, EmbeddingGemma, Qwen3-Embedding).
   - Embeddings hentes fra den aktive endpoint med den valgte embedding-model.
   - Skiftes model eller endpoint, bygges indekset forfra, fordi vektorer fra forskellige modeller ikke kan sammenlignes.
   - Understøtter endpointen ikke `/v1/embeddings`, vises en konkret fejl med koden `NO_EMBEDDINGS`.
5. **Holdes opdateret**:
   - Indeksering ved start, efter gem og på knappen *Indeksér nu*.
   - Derudover ved filændringer (`fs.watch` rekursivt, med 3 s debounce).
   - Kun ændrede filer (mtime/størrelse) indekseres igen, og slettede fjernes.
   - Skjulte filer, Office-låsefiler (`~$…`), `node_modules` og `.git` springes over, ligesom filer over 30 MB.
6. **Chat med kilder**:
   - Med *Søg i mine mapper* slået til i chatten søger bridgen med spørgsmålet (ved korte opfølgninger også den forrige besked) og giver modellen de bedste uddrag som nummererede `<source>`-blokke. Standard er 5 uddrag, og lighed under 0,2 frasorteres.
   - Uddragene er data, ikke instruktioner, og modellen bedes citere med [1], [2] …
   - Bridgen sender en `sources`-hændelse, og UI'et viser kilderne under svaret med uddrag og *Åbn fil*.
   - Kilder kommer altid fra bridgens eget indeks. Det, en klient selv sender, smides væk.
7. **Åbn fil**: Bridgen åbner filen i dens standardprogram (`open`, `explorer.exe`, `xdg-open` via `execFile`, uden shell). Kun filer, der står i indekset: stien slås op fra chunk-id'et og tages aldrig fra forespørgslen.
8. **Politik og audit**:
   - `policy.knowledge: false` slår funktionen fra.
   - `policy.knowledgeFolders` begrænser til bestemte mapper og alt under dem. Det håndhæves ved gem og ved hver indeksering; afviste mapper vises i status.
   - Audit-loggen får `index`- og `retrieve`-poster med host, model, varighed og antal. Aldrig stier, spørgsmål eller indhold.

## Konsekvenser

- Kun bridge og MIAI App. Direkte tilstand viser, at funktionen kræver bridgen.
- Mappestier skrives ind som tekst, fordi webviewet ikke har en mappevælger. MIAI App kan senere tilbyde en native dialog.
- Hele indekset ligger i hukommelsen under søgning: ca. 4 KB pr. uddrag med 1.024 dimensioner, altså ca. 400 MB ved 100.000 uddrag. Det er acceptabelt for målgruppen; se punkt 2 for vejen videre.
- Kvaliteten afhænger af embedding-modellen. Anbefalingerne står i PLAN.md §5 og BRUG.md.
