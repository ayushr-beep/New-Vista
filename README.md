# VISTA — Virventures Intelligent Shipment & Trade Analytics

Production-ready static web app. No backend, no server costs — runs entirely
in the browser and deploys free on GitHub Pages.

## Structure
```
vista/
├── index.html        # markup only (tabs, panels, layout)
├── css/styles.css     # all styling, themes (dark/light), animations
└── js/app.js          # data parsing, state, charts, tabs, auth, AI assistant stub
```

## Run it locally
Just open `index.html` in a browser — no build step, no install.
(For local testing with file uploads working cleanly, you can also run
`python3 -m http.server` in this folder and visit `localhost:8000`.)

## Deploy for free — GitHub Pages
1. Create a new GitHub repo, e.g. `vista`.
2. Push these three files/folders to the repo root.
3. Repo → Settings → Pages → Source: `main` branch, `/root`.
4. Your live URL: `https://<your-username>.github.io/vista/`

Every time you push a change, the live site updates automatically — free,
forever, no server to maintain.

## What changed from the original single-file version
- CSS and JS pulled out of inline `<style>`/`<script>` blocks into their own
  files — same functionality, but now git-diffable, cacheable by the browser,
  and easier to hand off/maintain.
- No logic was changed in this step. This is a structural pass only.

## Roadmap (next steps)
1. ✅ Modularize into index.html / css / js (this step)
2. Manifest tab: real min-split + regional-demand allocation solver
3. Cross-tab enrichment: surface SKU / vendor / velocity / regional /
   inventory context consistently across every tab
4. (Optional) Secure AI assistant backend via a free serverless proxy
