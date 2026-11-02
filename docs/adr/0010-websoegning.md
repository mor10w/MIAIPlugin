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
3. **Indstilling**: `webSearch` (boolean, standard: til) i fælles `Settings`.
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
