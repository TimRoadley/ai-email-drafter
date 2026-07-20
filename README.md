# AI Email Drafter

An Outlook add-in that uses your own OpenAI-compatible LLM endpoint (e.g. `https://your-llm-host.com/v1`) to:

- Draft emails
- Draft email replies based on the current email thread
- Improve writing

## Install

Outlook requires sideloading the add-in from a local manifest file (no repo clone needed):

1. Download the manifest: [manifest.xml](https://timroadley.github.io/ai-email-drafter/manifest.xml) (right-click → Save Link As)
2. In Outlook: **Preferences → Extensions → Add Custom Add-in → Add from file...**
3. Select the downloaded `manifest.xml`

The add-in itself is loaded from GitHub Pages, so the manifest file just needs to exist on disk at install time — you can delete it afterward. It persists in Outlook across restarts; you only need to do this once.

## Configure

After installing, click the ⚙ **Settings** button in the taskpane:

- **LLM Base URL** — your OpenAI-compatible endpoint. Must be an absolute `http(s)://` URL.
- **API Key** (optional) — sent as an `Authorization: Bearer <key>` header. Leave blank if your endpoint doesn't require one.

Settings are saved per-user in Outlook roaming settings.

## Troubleshooting

- **"AI Offline"** — verify the Base URL is reachable (`curl <baseUrl>/models` should return JSON). Enter an API key in Settings if the endpoint requires one.
- **401/403 from the LLM** — enter the API key in Settings.
- **"Add-in could not be loaded"** — the deployment may be out of date; try again shortly or contact the maintainer.
- **Changes not reflecting** — close/reopen the compose window, toggle the add-in off/on, or restart Outlook (it caches aggressively).

## Removing the add-in

Outlook → Preferences → Extensions → My Add-ins → remove the entry.

## Contributing
See [CONTRIBUTE.md](CONTRIBUTE.md)
