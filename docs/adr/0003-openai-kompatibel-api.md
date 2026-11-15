# ADR-0003: Udelukkende OpenAI-kompatibel API mod inference

- Status: Vedtaget
- Dato: 2026-11-02

## Kontekst
Brugeren skal kunne bruge et hvilket som helst lokalt inference-værktøj. Ollama, LM Studio, llama.cpp-server, vLLM, LocalAI, Jan, KoboldCpp, text-generation-webui, TGI og Foundry Local udstiller alle `/v1/chat/completions` og (de fleste) `/v1/models`.

## Beslutning
Bridgen har én provider: en generisk OpenAI Chat Completions-klient. En endpoint består af `baseUrl`, valgfri `apiKey` og capability-flag:

- `structuredOutput`: `json_schema` | `json_object` | `none`
- `stripThinking`: fjern `<think>`-blokke (standard: til)

Ingen leverandørspecifikke adaptere. Cloud-udbydere understøttes ikke direkte; vil en organisation det, kan den selv stille en OpenAI-kompatibel gateway op, som politikken tillader.

## Konsekvenser
+ Nye værktøjer virker uden kodeændringer.
+ Lille, testbar kodebase.
- Forskelle i struktureret output og tool calling håndteres med flag og fallback-parsing.
- Modelstyring (download/pull) ligger i værktøjet, ikke i MIAI.

## Tillæg: tænkning og indsats (2026-11)

Brugeren kan slå modellens tænkning til eller fra og vælge en indsats (Indstillinger → Svar). Der er ingen fælles standard for det på tværs af værktøjerne, så MIAI sender valget i flere dialekter på én gang:

| Felt | Forstås af | Fra | Til | Indsats |
|---|---|---|---|---|
| `reasoning_effort` | OpenAI, Ollama, LM Studio (0.4.8+), vLLM for nogle modeller | `"none"` | (udelades) | `"low"`, `"medium"`, `"high"` |
| `think` | Ollama | `false` | `true` (kun uden indsats) | – |
| `chat_template_kwargs.enable_thinking` | vLLM og llama.cpp (`--jinja`) med Qwen3 m.fl. | `false` | `true` | – |

- **Serverens standard** (udgangspunktet) sender ingen af felterne, så intet ændrer sig for eksisterende opsætninger.
- `chat_template_kwargs` flettes med det, brugeren selv har lagt i endpointens `extraBody`.
- **Tilbagefald:** de fleste servere ignorerer ukendte felter, men strenge (OpenAI, nogle gateways) svarer 400. Nævner fejlen tænkning (`think`, `reasoning`, `chat_template`), prøver MIAI igen uden udvidelserne og derefter helt uden. Det huskes pr. server og model i sessionen, ligesom `stream_options`. Andre 400-fejl prøves ikke igen.
- Sendes forespørgslen helt uden valget, får svaret advarslen `thinking-ignored`, som task pane viser under svaret.
- Den tænkning, modellen eventuelt laver, filtreres stadig fra svaret (`<think>` og `reasoning_content`).

**Konsekvenser:**

- Virker kun med modeller, der kan styres. Rene tænke-modeller (fx Qwen3-Thinking) tænker altid, og gpt-oss kan ikke slås helt fra.
- Ikke afprøvet mod rigtige servere ud over den simulerede i testene; feltnavnene bygger på værktøjernes dokumentation (Ollama, vLLM, LM Studio, november 2026).

