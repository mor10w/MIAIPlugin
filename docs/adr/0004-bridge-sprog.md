# ADR-0004: Bridge implementeres i Node.js/TypeScript

- Status: Vedtaget
- Dato: 2026-11-02

## Kontekst
Muligheder var Node.js/TypeScript, .NET 8 og Rust (Tauri).

## Beslutning
Node.js 22 LTS + TypeScript + Fastify 5. Typer og opgavedefinitioner deles med add-in'et via pakken `@miai/shared`.

## Konsekvenser
+ Ét sprog, delte typer og validering (Zod) på begge sider.
+ Fastify har god SSE- og statisk fil-understøttelse.
- Distribution kræver en pakket Node-binær (Node SEA eller tilsvarende); afklares i fase 1.
- Windows-service og cert-store håndteres af installeren frem for af bridgen.
