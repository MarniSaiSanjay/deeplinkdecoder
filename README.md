# Deep Link Decoder

A small, self-contained web app that decodes **Microsoft 365 Copilot** and **Word**
deep links into their component parts — the base64 payload, the substrate
identifier (encoding, endpoint, prompt ID, locale) and the query-string
parameters.

It is the inverse of the link generator: paste a link, see exactly what it
encodes. All decoding happens **in the browser** — nothing is sent to a server.

**Live demo:** https://marnisaisanjay.github.io/deeplinkdecoder/

## What it decodes

- **M365 Copilot chat (prompt)** — `https://m365.cloud.microsoft/chat/entity1-<id>/<payload>?...`
  - `MicrosoftV2` identifiers → endpoint, link GUID, prompt ID, locale
  - `MicrosoftHandOffV5` identifiers (Researcher with prompt) → prompt ID, locale
- **Word handoff (growth)** — `https://word.cloud.microsoft/handoff/growth/?handinPayload=<payload>`
  - v1.1 (nested base64) and v1.2 (`promptId_locale`) schemas
- **Base / Create / Store / Researcher** links — surfaced with their query params.

## Tech stack

- React 19 + TypeScript
- Vite

## Local development

```bash
npm install
npm run dev      # http://localhost:5173/DeepLinkDecoder/
```

## Project structure

```
src/
  decoder.ts      Core decoding logic (the inverse of the link generator)
  App.tsx         UI: input box + rendered breakdown
  main.tsx        React entry point + global style reset
  theme.ts        Shared design tokens
index.html        Vite HTML entry
vite.config.ts    Build config (GitHub Pages base path)
.github/workflows/deploy.yml   CI build + Pages deploy
```

## Build

```bash
npm run build    # outputs static site to ./dist
npm run preview  # preview the production build locally
```

## Deploy to GitHub Pages

A workflow at `.github/workflows/deploy.yml` builds and publishes to GitHub Pages
on every push to `main`.

1. Push this project to a GitHub repository.
2. In **Settings → Pages**, set **Source** to **GitHub Actions**.
3. Push to `main`. The site publishes at `https://<user>.github.io/<repo>/`.

The Vite `base` path is derived automatically from the repository name in CI.
For a local production build under a sub-path, run:

```bash
VITE_BASE=/your-repo-name/ npm run build
```
