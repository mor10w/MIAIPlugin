# Arbejdsgang for MIAIPlugin

- **Ingen direkte commits på `main`.** Al ændring laves på en gren og merges via en pull request.
  - Gren-navn: `<type>/<kort-beskrivelse>`, fx `fix/mermaid-etiketter` eller `feat/excel-pivot`.
  - Én pull request pr. sammenhængende ændring, med en kort beskrivelse af hvad og hvorfor og hvordan det er testet.
  - Merge først, når CI er grøn (lint, e2e, build på Ubuntu og Windows).
- Kommunikation, commit-beskeder og dokumentation skrives på dansk.
- Før push: `npx prettier --check .`, `npm run lint -- --max-warnings 0`, `npm test` og de berørte e2e-tests (`npx playwright test -c packages/addin/e2e/playwright.config.ts <fil>` efter `npm run build`). De visuelle tests (`visual.spec.ts`) kan kun sammenlignes i CI.
