# Contributing

## Project structure

| Path | Purpose |
|------|---------|
| `src/teams-export-prototype.js` | Main export tool — self-contained IIFE injected into Teams |
| `extension-src/` | Chrome MV3 extension source (manifest, background script, icons) |
| `extension-dist/` | Built extension output (gitignored) |
| `scripts/` | Build, packaging, validation, and probe automation |
| `tests/` | Unit and integration tests |
| `docs/findings.md` | Live DOM notes and validation log |
| `artifacts/` | Validation outputs, screenshots, exports (gitignored) |

## Prerequisites

- Node.js (latest LTS)
- Chrome or Chromium

## Setup

```sh
npm install
```

## Scripts

| Command | Description |
|---------|-------------|
| `npm run build:extension` | Build unpacked extension into `extension-dist/` |
| `npm run package:extension` | Build and zip extension to `artifacts/teams-selected-messages-export-extension.zip` |
| `npm test` | Run unit and integration tests |
| `npm run test:unit` | Unit tests only |
| `npm run test:integration` | Integration tests (requires Chrome/Chromium) |
| `npm run test:e2e` | End-to-end validation against live Teams (requires signed-in Chrome debug session) |
| `npm run chrome:launch` | Launch a dedicated Chrome debug instance |
| `npm run chrome:launch:extension` | Launch Chrome with extension preloaded on debug port `9223` |
| `npm run extension:load` | Load unpacked extension into a running Chrome debug instance |
| `npm run validate:live` | Validate the injected prototype against live Teams |
| `npm run validate:extension` | Validate the installed extension against live Teams |
| `npm run probe:worker` | Run the experimental worker interception probe |

## Architecture

The export tool is a single self-contained IIFE (`src/teams-export-prototype.js`) that:

1. Injects export controls into the Teams title bar
2. Manages message selection state via DOM event listeners
3. Scrapes message content, replies, reactions, and metadata from the Teams DOM
4. Exports to Markdown or HTML with formatted output

When packaged as a Chrome extension, the content script is a copy of this IIFE injected at `document_idle` on Teams pages, with a background service worker handling the extension action click.

### Injection mode (development)

For rapid iteration, paste or inject `src/teams-export-prototype.js` directly into the Teams DevTools console. The script exposes `window.__teamsMessageExporter` for debugging.

### Validation harnesses

The project includes two CDP-based validation harnesses:

- **Injected validation** (`scripts/validate-live-prototype.mjs`) — connects over DevTools protocol, injects the prototype, and validates selection + export
- **Extension validation** (`scripts/validate-installed-extension.mjs`) — uses real pointer clicks against the installed extension UI

Both write artifacts to `artifacts/exports/` and `artifacts/screenshots/`.

### Worker interception probe

`scripts/probe-worker-intercepts.mjs` is an experimental probe that patches `window.Worker` at page start to capture Teams GraphQL operations flowing through the precompiled web worker. This is documented in `docs/findings.md` under "Worker Hook Next Steps".

## Testing

- **Unit tests** (`tests/unit/`) — pure function tests for worker-probe summarization helpers
- **Integration tests** (`tests/integration/`) — run the export UI against a local Teams-like HTML fixture in headless Chrome via Puppeteer
- **E2E tests** — reuse the extension validation harness against a live signed-in Teams session

## Release

Pushing to the `main` branch triggers a GitHub Actions workflow that builds, packages, and publishes the extension zip as a GitHub Release. The version is read from `extension-src/manifest.json`.
