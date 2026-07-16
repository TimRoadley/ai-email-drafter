# Contributing to AI Email Drafter

Thanks for helping improve this project. This guide covers dev setup, local debugging, and how to ship changes.

## Prerequisites

- macOS 13+
- Node.js 18+
- Outlook
- Access to an OpenAI-compatible LLM endpoint for testing

## Setup

```bash
npm install
npm start
```

This starts a hot-reloading HTTPS dev server on `https://localhost:3000`.

If Outlook rejects the self-signed cert, run `npx office-addin-dev-certs install` (or accept it in Safari). Stop the dev server with `Ctrl+C`.

## Debugging in Outlook locally

1. Temporarily edit the URLs in `manifest.xml` and `manifest.json` (repo root) to point at `https://localhost:3000` instead of the GitHub Pages URL (also update `validDomains` / `AppDomain` to `localhost`).
2. Build and sideload once:
   ```bash
   npm run build:dev
   # sideload dist/manifest.xml in Outlook (Preferences → Extensions → Add Custom Add-in → Add from file...)
   ```
3. **Revert the manifest changes before committing/pushing** — never commit localhost URLs.

## Linting & validation

```bash
npm run lint          # check for lint errors
npm run lint:fix      # fix auto-fixable issues
npm run validate      # validate dist/manifest.xml against the Office schema (requires a build first)
```

Run lint before opening a PR.

## Build scripts

| Script | Purpose |
|---|---|
| `npm run build` | Production webpack build (bundles everything into `dist/`) |
| `npm run build:dev` | Webpack build in dev mode (bundles everything into `dist/`) |
| `npm run validate` | Validate the built `dist/manifest.xml` against the Office add-in schema |
| `npm run lint` | Lint via `office-addin-lint` |
| `npm start` | Hot-reloading dev server on `https://localhost:3000` |

## Architecture notes

- No server-side component — the add-in runs entirely client-side against a user-supplied LLM endpoint.
- `manifest.xml` and `manifest.json` in the repo root are the **production** manifests, pointing at `https://timroadley.github.io/ai-email-drafter/`. GitHub Pages serves the *contents* of `dist/` at the site root, so deployed URLs do **not** include a `dist/` segment — only local build output paths do. Only edit the manifest URLs for local debugging (and revert before pushing) or for real production changes.
- The LLM endpoint must serve the OpenAI-compatible API (`/models`, `/chat/completions`). Users configure the Base URL and API key themselves in Settings.
- Source layout: `src/taskpane/` (UI, ai-service, settings, prompt template), `src/commands/` (ribbon command entry point).

## Submitting changes

1. Branch from `main`, make your changes.
2. Run `npm run lint` and `npm run build` to confirm everything builds cleanly.
3. Open a PR against `main`.

## Releasing (maintainers)

Deployment is automated via GitHub Actions (`.github/workflows/deploy.yml`). Merging/pushing to the `prod` branch triggers a build and deploy to GitHub Pages:

```bash
git checkout prod
git merge main
git push origin prod
```

The workflow installs Node 20, runs `npm ci && npm run build`, and deploys `dist/` to GitHub Pages. No re-sideload is needed for end users unless the Pages URL itself changes.
