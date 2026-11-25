# ADR-0010: Web-søgning for modellen (tool-calling via bridgen)

- **Status:** Beslut
- **Dato:** 2026-11-02

## Kontekst

Modellen kører lokalt (oMLX) og har derfor intet internetforbindelse – den svarer kun
ud fra træningsdata, og dens viden er forældet. Brugeren vil have modellen i stand til
at søge online efter svar og inspiration, og at den er *klar over* muligheden, så det
selv beslutter hvornår en søgning er relevant.

Verifikation af forudsætninger (testet mod oMLX 2026-11-02):
- oMLX understøtter OpenAI's `tools`-parameter og strømmer `tool_calls`-deltas korrekt.
- Modellen (Qwen3 27b) kalder værktøjet når det er relevant, svarer direkte hvis det
  kan, og forfølger et svagt søgeresultat med en mere præcis opfølgende søgning.

Søgningen skal ske i **bridgen** (Node), ikke i browseren: DuckDuckGo's HTML-endpoint
har ingen CORS-støtte, og en eventuel API-nøgle (fx Brave) må aldrig sendes til
add-in'ets webview.

## Beslutning

1. **Tool-calling-løkke i den fælles motor** (`shared/engine.ts`): Motoren får et
   valgfrit `tools`-afhængighed (specifikationer + en `run()`-eksekverer). Når det er
   leveret og web-søgning er slået til i indstillingerne, tilføjer motoren en linje om
   muligheden til systemprompten og kører indtil modellen stopper – maks. 4 runder,
   hvor hver rundes værktøjsresultater sendes tilbage som `tool`-beskeder.
2. **Søgebackend uden API-nøgle**: DuckDuckGo's HTML-endpoint (`html.duckduckgo.com`)
   skraperes i `bridge/search.ts`. Resultater (titel, URL, snippet) formateres som
   tekst til modellen. Ingen nye afhængigheder (ren `fetch` + regex-parsning).
3. **Indstilling**: `webSearch` (boolean, **standard: fra** – rettet ved code review 2026-11-03) i fælles `Settings`.
   Bridgen leverer værktøjet kun når det er til. I direkte tilstand (browser → model
   uden bridge) leveres intet værktøj – funktionaliteten er således automatisk
   bridge-baseret, som architecturen kræver.
4. **Gennemsigtighed**: Bridgen sender en `search`-event med søgeforespørgslen, så
   task pane kan vise "Søger på web…". Modellen opfordres i prompten til at henvise
   til kilder, når den bruger søgeresultater.

## Konsekvenser

- **Fordele:** Modellen kan besvare aktuelle spørgsmål (nyheder, versioner, priser) og
  bruge webbet som inspiration – uden at brugeren selv skal søge og kopiere.
- **Sikkerhed:** Søgeresultater er utroværdig input (prompt injection kan gemmes i
  websider). De indsendes som `tool`-beskeder, og systemprompten fastholder
  `<document>`/kontekst-reglerne; modeloutput renses stadig som før.
- **Begrænsninger:** DuckDuckGo kan rate-limitere ved meget høj belastning (svaret
  bliver da "søgetjenesten svarede ikke"); resultater er tekst-snippets, ikke fulde
  sider. En API-nøgle-baseret provider (fx Brave) kan tilføjes bag samme
   `run()`-grænse, hvis det bliver nødvendigt.
- **Overvågning:** Hver opgave kan køre op til 4 model-runder; audit-loggen fortsætter
  med at logge kun metadata.

## Tillæg (code review 2026-11-03)

Søgeordene dannes af modellen ud fra dokumentet og sendes til en ekstern tjeneste. Det er den eneste
udgående trafik, der ikke går til den konfigurerede model-endpoint, og den var hverken omfattet af
politikken (ADR-0006) eller audit-loggen. Rettet:

- `Settings.webSearch` er **fra** som standard (aktivt valg, jf. løftet om at data bliver lokalt).
- Ny politik-indstilling `Policy.webSearch` (standard `true`). En organisation kan sætte den til `false`
  i maskinpolitikken og dermed forbyde søgning; motoren kræver både indstilling og politik.
- Hver søgning skrives i audit-loggen (`kind: "search"`, host `html.duckduckgo.com`, varighed – aldrig søgeord).
- Værktøjskald sendes tilbage med `type: "function"`, som OpenAI-specifikationen kræver (vLLM m.fl. afviser ellers).
- Indstillingen vises kun i bridge-tilstand og er låst, når politikken forbyder søgning.

## Tillæg (#12, 2026-11-03): selvhostet søgetjeneste

- **Indstilling:** `Settings.search = { provider: "duckduckgo" | "searxng", url }`. SearXNG bruges via JSON-API'et (`/search?format=json`).
- **Politik:** adressen valideres mod `allowedHosts` både ved gem og ved hver søgning. En maskinpolitik kan låse tjenesten med `Policy.search`.
- **Audit-log:** registrerer den faktiske søgehost.
- **DuckDuckGo:** en side med resultater, som ikke kan læses (ændret markup), skelnes nu fra "ingen resultater". Modellen og loggen får besked, så et brud opdages.

## Tillæg (2026-11-25): Exa som søgetjeneste, og robot-tjek hos DuckDuckGo

Afprøvningen af `=MIAI.ASK(…; FALSK)` viste to svagheder ved DuckDuckGo som eneste tjeneste uden opsætning:

- DuckDuckGo svarer med et robot-tjek (HTTP 202) efter få søgninger på kort tid. En celle kan søge op til fire gange, så en kolonne rammer det hurtigt. Bridgen meldte det som "ingen resultater"; nu meldes søgningen som mislykket, og modellen får besked om ikke at gætte.
- Resultaterne er korte uddrag. Et præcist tal, fx en lukkekurs på en bestemt dag, står sjældent i dem.

Derfor kan `Settings.search.provider` nu også være `"exa"`:

- **Hvordan:** bridgen kalder værktøjet `web_search_exa` på Exas hostede MCP-server (`https://mcp.exa.ai/mcp`) med ét JSON-RPC-kald. Serveren kræver hverken session, konto eller nøgle, så beslutningen om søgning uden API-nøgle holder. Svaret er tekst med titel, URL og uddrag af selve siden; det afkortes til 8.000 tegn, og der bedes om fire resultater.
- **Ikke standard:** DuckDuckGo er stadig standard. Exa er et aktivt valg, fordi søgeordene går til en anden tredjepart, og fordi Exas vilkår for brug af den nøglefri adgang fra et produkt som MIAI ikke er afklaret.
- **Grænser:** uden nøgle gælder Exas gratis grænse (intet tal er oplyst). Er den nået, svarer serveren 429; modellen får samme besked som ved robot-tjekket. Nøglefri adgang kræver ifølge Exa en direkte IPv4-forbindelse.
- **Politik og audit-log:** som DuckDuckGo er Exa en fast, offentlig adresse og tjekkes ikke mod `allowedHosts`. `Policy.search` kan låse tjenesten til `{"provider": "exa"}`, og `Policy.webSearch: false` forbyder al søgning. Audit-loggen får hosten `mcp.exa.ai`.
- **Ikke med:** en valgfri Exa-nøgle for højere grænser (kræver samme håndtering i nøgleringen som endpoint-nøgler) og Exas `web_fetch_exa`, som henter en bestemt side.
- **Dags dato:** når modellen kan søge, får systemprompten dags dato. Uden den afviste modellen en dato i sidste uge som "i fremtiden".
