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
