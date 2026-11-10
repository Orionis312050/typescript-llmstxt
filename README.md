# typescript-llmstxt

Édition **offline** de la doc TypeScript du Teams SDK, à partir de https://microsoft.github.io/teams-sdk/llms_docs/llms_typescript.txt (+ contenu complet).

## Structure

- `llms.txt` : index de toutes les pages (liens relatifs vers `docs/`).
- `docs/<categorie>/<feature>.md` : **une feature par fichier**, contenu complet inline.
- Chaque dossier de catégorie contient un `_category_.json` (compatible Docusaurus).
- Tous les cross-references sont des liens relatifs intra-dépôt : aucun fetch externe requis.

## Source

Généré depuis `microsoft/teams-sdk` (`teams.md/src/pages/templates` + fragments `typescript.incl.md`, et `teams.md/docs/main`), même pipeline que `generate-llms-txt.ts`.
