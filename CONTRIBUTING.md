# Contributing

## Project structure

| Path | Purpose |
|------|---------|
| `src/` | TypeScript source modules (entry point: `src/main.ts`) |
| `dist/` | Built content script IIFE bundle (gitignored) |
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
| `npm run build` | Bundle TypeScript source into `dist/content-script.js` |
| `npm run typecheck` | Run TypeScript type checking without emitting |
| `npm run build:extension` | Build bundle and copy to unpacked extension in `extension-dist/` |
| `npm run package:extension` | Build, copy, and zip extension to `artifacts/teams-selected-messages-export-extension.zip` |
| `npm test` | Build, then run unit and integration tests |
| `npm run test:unit` | Unit tests only (TypeScript and MJS) |
| `npm run test:integration` | Integration tests (requires Chrome/Chromium and a prior build) |
| `npm run test:e2e` | End-to-end validation against live Teams (requires signed-in Chrome debug session) |
| `npm run chrome:launch` | Launch a dedicated Chrome debug instance |
| `npm run chrome:launch:extension` | Launch Chrome with extension preloaded on debug port `9223` |
| `npm run extension:load` | Load unpacked extension into a running Chrome debug instance |
| `npm run validate:live` | Validate the injected prototype against live Teams |
| `npm run validate:extension` | Validate the installed extension against live Teams |
| `npm run probe:worker` | Run the experimental worker interception probe |

## Architecture

The export tool is built from modular TypeScript source files in `src/`, bundled by esbuild into a single IIFE (`dist/content-script.js`) that:

1. Injects export controls into the Teams title bar
2. Manages message selection state via DOM event listeners
3. Scrapes message content, replies, reactions, and metadata from the Teams DOM
4. Exports to Markdown or HTML with formatted output

### Source modules

| Module | Responsibility |
|--------|---------------|
| `main.ts` | Entry point — wires callbacks, orchestrates startup |
| `types.ts` | All TypeScript interfaces |
| `constants.ts` | CSS class names, DOM selectors, timing values |
| `state.ts` | Singleton state object and callback registry |
| `utilities.ts` | Pure utility functions and DOM predicates |
| `styles.ts` | CSS stylesheet generation |
| `theme.ts` | Teams light/dark theme detection |
| `conversation.ts` | Conversation title/key detection from DOM |
| `content-extraction.ts` | DOM content extraction (mentions, images, quotes, reactions) |
| `strategy.ts` | DOM strategy scoring and message row detection |
| `messages.ts` | Message record building from DOM elements |
| `selection.ts` | Pure selection helpers (range, ordering) |
| `markdown-renderer.ts` | Markdown conversion and document rendering |
| `html-renderer.ts` | HTML document rendering for export |
| `export-helpers.ts` | Export payload creation, download triggers, clipboard |
| `history.ts` | Full chat history scroll harvesting |
| `toolbar.ts` | Toolbar DOM creation, dock positioning, busy state |
| `lifecycle.ts` | Global event handlers, MutationObserver, extension bridge |

When packaged as a Chrome extension, the bundled content script is injected at `document_idle` on Teams pages, with a background service worker handling the extension action click.

### Injection mode (development)

For rapid iteration, build with `npm run build` and then inject `dist/content-script.js` into the Teams DevTools console. The script exposes `window.__teamsMessageExporter` for debugging.

### Validation harnesses

The project includes two CDP-based validation harnesses:

- **Injected validation** (`scripts/validate-live-prototype.mjs`) — connects over DevTools protocol, injects the prototype, and validates selection + export
- **Extension validation** (`scripts/validate-installed-extension.mjs`) — uses real pointer clicks against the installed extension UI

Both write artifacts to `artifacts/exports/` and `artifacts/screenshots/`.

### Worker interception probe

`scripts/probe-worker-intercepts.mjs` is an experimental probe that patches `window.Worker` at page start to capture Teams GraphQL operations flowing through the precompiled web worker. This is documented in `docs/findings.md` under "Worker Hook Next Steps".

## Testing

- **Unit tests** (`tests/unit/`) — TypeScript tests for core functions (utilities, selection, renderers, export helpers) and MJS tests for worker-probe summarization helpers
- **Integration tests** (`tests/integration/`) — run the export UI against a local Teams-like HTML fixture in headless Chrome via Puppeteer
- **E2E tests** — reuse the extension validation harness against a live signed-in Teams session

## Release

Pushing to the `main` branch triggers a GitHub Actions workflow that builds, packages, and publishes the extension zip as a GitHub Release. The version is read from `extension-src/manifest.json`.
