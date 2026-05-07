# Homebase

Personal dashboard PWA — weather, Outlook mail and calendar, live sports
ticker, news headlines, rich-text notes with Apple Pencil handwriting,
to-do list, and a Claude-powered AI assistant. Notes and to-dos sync
across devices via your OneDrive when signed in to Microsoft.

Built as a static site (no backend). All API keys live in your browser's
localStorage; requests go directly from your device to the third-party
services.

## Run locally

Any static server works. From this directory:

```sh
python3 -m http.server 8000
# then open http://localhost:8000
```

## Deploy to Render

`render.yaml` is a static-site blueprint. Connect this repo in Render
and it will deploy automatically. PR previews are enabled.

## Configuration

Open the dashboard, click ⚙ Settings, and add the keys you want:

| Service | Required for | Where to get a key |
| --- | --- | --- |
| Microsoft Azure App Client ID | Outlook mail/calendar, cross-device sync of notes/to-dos | portal.azure.com → App registrations |
| Anthropic API key | AI assistant | console.anthropic.com → API keys |
| NewsAPI key (optional) | Branded news headlines | newsapi.org (falls back to Hacker News if blank) |

### Microsoft Azure setup

1. portal.azure.com → **App registrations** → **New registration**
2. Platform: **Single-page application (SPA)**
3. Redirect URI: the URL where this dashboard is hosted (e.g.
   `http://localhost:8000/` for local, your Render URL for production).
   You can add multiple — add one per environment.
4. API permissions (delegated, all under Microsoft Graph):
   - `User.Read`
   - `Mail.Read`
   - `Calendars.Read`
   - `Files.ReadWrite.AppFolder`
   - `Tasks.ReadWrite`
5. Copy the **Application (client) ID** into Settings.

### News

NewsAPI's free Developer plan only allows browser requests from
`localhost`, so deployed sites should leave the key blank (Hacker News
fallback works without configuration) or upgrade to a paid plan.

## File layout

```
index.html            entry / PWA shell
dashboard.css         theme + layout
dashboard.jsx         the app (React via Babel-standalone)
manifest.webmanifest  PWA manifest
sw.js                 service worker
icons/                app icons
render.yaml           Render static-site config
```
