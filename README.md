# 📋 Pastoral Social — Cadastro — installable app (PWA wrapper)

This branch holds the small static site that lets volunteers **install the
[Cadastro-PS-PSB](https://github.com/kauemelo01/Cadastro-PS-PSB) Streamlit app on their phones**,
with its own name and icon, opening full screen like a native app.

- **Install page:** https://kauemelo01.github.io/Cadastro-PS-PSB/
- **Wrapped app:** https://cadastro-ps-psb.streamlit.app/
- **Installed name:** "Cadastro" (clipboard icon)

It is published by GitHub Pages straight from this `gh-pages` branch (root folder).
The app's code lives on `main`; nothing here touches it.

## Why a wrapper?

Streamlit Community Cloud controls the app's HTML page, so the app can't ship its own
manifest or service worker. Installed straight from `streamlit.app`, it shows up as
"Streamlit" and offers no install button. This site supplies the missing pieces and shows
the app full-screen inside an `<iframe>`.

## Files

| File | Purpose |
|---|---|
| `index.html` | The page: install bar on top, the app in a full-screen `<iframe>` below |
| `manifest.json` | App name, icons, colours, `"display": "fullscreen"` |
| `sw.js` | Service worker: makes the site installable and caches the wrapper's own files (never the app) |
| `icon-192.png`, `icon-512.png` | Android icons (`icon-512.png` is also the maskable icon) |
| `apple-touch-icon.png` | iPhone/iPad home-screen icon (180×180) |
| `.nojekyll` | Tells GitHub Pages to serve the files as-is |

## How it behaves

**Install bar** (green bar at the top, in Portuguese):

| Device | What the user sees |
|---|---|
| Android (Chrome) | **📲 Instalar app** opens Chrome's install dialog directly |
| iPhone / iPad (Safari) | **📲 Instalar app** shows the steps: Compartilhar → Adicionar à Tela de Início (iOS doesn't allow installing from code) |
| Computer | No bar |
| Already installed | No bar |

Tapping **✕** hides the bar for good on that device (stored in `localStorage`).

**Display**
- **Android:** the installed app opens full screen, with the status bar hidden.
- **iPhone:** opens as a normal app window. Safari ignores full-screen mode.

**Light / dark mode:** the page background follows the phone's setting: white, or
`#0e1117` (Streamlit's dark background) so no frame shows around the app. The splash
screen colour comes from `manifest.json` and is always white (a manifest allows only
one colour).

**Updates:** the service worker is network-first, so changes pushed here reach installed
apps the next time they open with internet. Chrome picks up `manifest.json` changes
(name, icons, display mode) on its own within about a day, or right away on reinstall.

## The iframe URL

The iframe loads `https://cadastro-ps-psb.streamlit.app/~/+/?embed=true`:

- **`/~/+/`** is the app itself. The plain `streamlit.app` address loads Streamlit Cloud's
  own outer page, which then loads the app in another frame; that outer page is always
  white and showed as a white frame around the app in dark mode.
- **`?embed=true`** hides Streamlit's toolbar and footer.

**Trade-off:** Streamlit Cloud's "wake this app up" button lives on that outer page. If
the app has gone to sleep from inactivity, the wrapper may show an error instead. To fix
it, open https://cadastro-ps-psb.streamlit.app/ once in a browser and tap the wake-up button; the
installed app then works again.

## Editing

1. **Get the branch:** check it out, or use a worktree so `main` stays untouched:
   ```bash
   git worktree add ../Cadastro-PS-PSB-pages gh-pages
   ```
2. **Make the change** and push to `gh-pages`. Pages republishes in about a minute.
3. **Changing the app address?** Edit the `<iframe src>` in `index.html`, keeping
   `/~/+/?embed=true` at the end.
4. **Adding or renaming cached files?** Update the `SHELL` list in `sw.js` and bump its
   `VERSION`, so old caches are cleared.
