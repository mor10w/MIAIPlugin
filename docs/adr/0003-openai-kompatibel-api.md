# ADR-0003: OpenAI-kompatibel API som fællesnævner for lokale modeller

- Status: Foreslået
- Dato: 2026-11-02

## Kontekst
Ollama, LM Studio, llama.cpp-server, vLLM og Foundry Local udstiller alle en OpenAI-kompatibel `/v1/chat/completions`.

## Beslutning
`OpenAICompatibleProvider` er standardadapteren. Native adaptere kun hvor de giver ekstra værdi (Ollama: modelstyring/`format`; Anthropic: Messages API; Bedrock: Converse API).

## Konsekvenser
+ Nye lokale runtimes understøttes ved konfiguration.
- Forskelle i tool calling/structured output skal håndteres med capability-flag pr. model.
