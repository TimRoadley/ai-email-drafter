# AI Email Drafter

This is an outlook add-in that uses your own OpenAI-compatible LLM endpoint (e.g. `https://your-llm-host.com/v1`) to:

  - Draft emails
  - Draft email replies
  - Improve writing

# Installation


---

## Production — deploy to GitHub Pages

**Prerequisite:** Node.js 18+.

### 1. Build the static bundle

```bash
npm install
npm run build
```

`dist/` now contains the taskpane, commands, office-js, icons, AND the `manifest.xml`/`manifest.json` files pointing at your GitHub Pages URL.

### 2. Publish to GitHub Pages

Commit the built `dist/` and push to the `main` branch. Enable GitHub Pages (repo **Settings → Pages**, source = `main` branch, `/root` folder) — or whatever branch/path your repo is configured to serve. The site must be served over HTTPS (GitHub Pages does this by default), since Outlook requires HTTPS for all add-in URLs.

### 3. Sideload the add-in once per user

On each user's Outlook:

1. **Outlook → Preferences → Extensions → Add Custom Add-in → Add from file...**
2. Provide the URL `https://timroadley.github.io/ai-email-drafter/dist/manifest.xml` (or download the built `dist/manifest.xml` and select it).

The manifest persists across Outlook restarts. Users do **not** need access to the repo.

---

## Configuration (per user, in the add-in)

After sideloading, click the ⚙ **Settings** button in the taskpane:

- **LLM Base URL** — your OpenAI-compatible endpoint (default: `https://api.example.com`). Must be an absolute `http(s)://` URL. Saved in Outlook roaming settings.
- **API Key** (optional) — sent to the endpoint as an `Authorization: Bearer <key>` header. Masked in the UI with a show/hide toggle. Leave blank if your endpoint doesn't require authentication. Stored in Outlook roaming settings.

---

## Dev

**Prerequisites:** macOS 13+, Node.js 18+, Outlook for Mac, and an OpenAI-compatible LLM endpoint.

```bash
npm install
npm start
```

Starts a hot-reloading HTTPS dev server on `https://localhost:3000`.

To debug against Outlook locally, temporarily edit the URLs in `manifest.xml` and `manifest.json` in the repo root to point at `https://localhost:3000` instead of the GitHub Pages URL. Then run:

```bash
npm run build:dev
# sideload dist/manifest.xml in Outlook once
```

**Do not commit these local changes** — revert them before pushing.

If Outlook rejects the self-signed cert, run `npx office-addin-dev-certs install` (or accept it in Safari). Stop the dev server with `Ctrl+C`.

### Linting & validation

```bash
npm run lint          # check for lint errors
npm run lint:fix      # fix auto-fixable issues
npm run validate      # validate dist/manifest.xml against the Office schema (requires built manifest)
```

---

## Build scripts

| Script | Purpose |
|---|---|
| `npm run build` | Production webpack build (bundles everything into `dist/`) |
| `npm run build:dev` | Webpack build in dev mode (bundles everything into `dist/`) |
| `npm run validate` | Validate the built `dist/manifest.xml` against the Office add-in schema |
| `npm run lint` | Lint via `office-addin-lint` |
| `npm start` | Hot-reloading dev server on `https://localhost:3000` |

---

## Architecture notes

- The add-in has **no server-side component**. The old `serve-local.js` / LaunchAgent / cert-generation machinery has been removed.
- `manifest.xml` and `manifest.json` in the repo root are the **production** manifests. They point at `https://timroadley.github.io/ai-email-drafter/dist`.
- For local development, temporarily edit the URLs in those files to `https://localhost:3000` (and update `validDomains` / `AppDomain` to `localhost`), build, sideload, then revert.
- The LLM endpoint must serve the OpenAI-compatible API (`/models`, `/chat/completions`). The default is `https://api.example.com`; users can change it in Settings.

---

## Updates & troubleshooting

To deploy a new version, rebuild, commit `dist/`, and push — users get the new bundle on next open. No re-sideload needed unless the Pages URL changed.

Removing the add-in: delete the entry in Outlook → Preferences → Extensions → My Add-ins.

- **"AI Offline"** — verify the Base URL is reachable (`curl <baseUrl>/models` should return JSON). Enter an API key in Settings if the endpoint requires one.
- **401/403 from the LLM** — enter the API key in Settings.
- **"Add-in could not be loaded" in Outlook** — the Pages URL in the manifest doesn't match the actual deployment URL, or the repo's Pages settings haven't published `dist/`.
- **Changes not reflecting in Outlook** — close/reopen the compose window, toggle the add-in off/on, or restart Outlook (it caches aggressively).
