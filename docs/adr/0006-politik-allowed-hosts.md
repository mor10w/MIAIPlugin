# ADR-0006: Politik for tilladte endpoints (allowedHosts)

- Status: Vedtaget
- Dato: 2026-11-02

## Kontekst
Løftet "data forbliver hos brugeren" skal kunne håndhæves, og senere skal en organisation kunne låse det centralt.

## Beslutning
Bridgen validerer hver endpoint-`baseUrl` mod en politik:

```json
{ "allowedHosts": ["127.0.0.1", "localhost", "::1"], "locked": false }
```

- Standard: kun loopback.
- Brugeren kan selv tilføje hosts (fx en GPU-maskine på LAN), medmindre `locked` er sat.
- En maskin-politikfil (`MIAI_POLICY_FILE`, senere fast placering under ProgramData / `/Library/Application Support`) har forrang og kan ikke ændres via UI.
- Wildcards: `*.firma.local`.

## Konsekvenser
+ Privatliv kan håndhæves og revideres.
+ Enterprise-udrulning kræver ingen kodeændringer.
- Brugere med remote-endpoint skal aktivt tilføje hosten.
