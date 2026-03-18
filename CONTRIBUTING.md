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
| `history.ts` | Full chat history export orchestration (API-first, scroll-harvest fallback) |
| `api-client.ts` | Direct REST API client — conversation ID resolution, message fetching and conversion |
| `worker-store.ts` | Bridge listener for MAIN-world worker hook — stores messages, members, auth tokens |
| `toolbar.ts` | Toolbar DOM creation, dock positioning, busy state |
| `lifecycle.ts` | Global event handlers, MutationObserver, extension bridge |

When packaged as a Chrome extension, the bundled content script is injected at `document_idle` on Teams pages. A separate `worker-hook.js` runs at `document_start` in the `MAIN` world to intercept worker traffic and capture auth tokens. The background service worker handles both extension icon clicks and cross-origin API fetch requests.

### Full chat export architecture

The full chat export uses a two-tier approach:

1. **Primary: REST API** — `worker-hook.js` reads the MSAL IC3 Bearer JWT from `localStorage` (keys matching `*accesstoken*ic3.teams.office.com*`) and the API region from the SKYPE-TOKEN discovery entry. These are bridged to the content script via `window.postMessage`, stored in `worker-store.ts`, and passed to `api-client.ts`. The background service worker (`background.js`) makes the actual cross-origin API calls to `{region}.ng.msg.teams.microsoft.com`, paginating through all messages.

2. **Fallback: Scroll-harvesting** — If the API path fails (no token, no conversation ID, API error), `history.ts` falls back to DOM scrolling: it scrolls the chat upward, captures visible message snapshots, and accumulates them until reaching the top.

### Build pipeline

The build has two steps that must both be run:

1. `node scripts/bundle.mjs` — esbuild bundles `src/main.ts` → `dist/content-script.js` (IIFE format)
2. `node scripts/build-extension.mjs` — copies `dist/content-script.js` + `extension-src/` files → `extension-dist/`

The `npm run build:extension` script runs both steps. Always use this when testing the extension locally.

### Cross-world communication

Teams Cloud uses a Service Worker that adds auth headers after the main-thread fetch, so intercepting fetch/XHR in the main world does not capture tokens. Instead:

- `worker-hook.js` runs in the **MAIN world** at `document_start` and reads tokens directly from `localStorage`
- It posts bridge messages via `window.postMessage({ source: "teams-export-worker-hook", ... })`
- `worker-store.ts` listens in the **ISOLATED world** and stores the received data
- DOM `dataset` attributes are used for cross-world data where needed (e.g., conversation ID)

### Injection mode (development)

For rapid iteration, build with `npm run build` and then inject `dist/content-script.js` into the Teams DevTools console. The script exposes `window.__teamsMessageExporter` for debugging.

### Validation harnesses

The project includes two CDP-based validation harnesses:

- **Injected validation** (`scripts/validate-live-prototype.mjs`) — connects over DevTools protocol, injects the prototype, and validates selection + export
- **Extension validation** (`scripts/validate-installed-extension.mjs`) — uses real pointer clicks against the installed extension UI

Both write artifacts to `artifacts/exports/` and `artifacts/screenshots/`.

### Worker interception

`worker-hook.js` also patches `window.Worker` at page start to capture Teams GraphQL operations flowing through the precompiled web worker. This provides structured message data, member identities, and reaction details that enrich DOM-extracted message records. The probe script at `scripts/probe-worker-intercepts.mjs` demonstrates standalone interception and is documented in `docs/findings.md`.

## Testing

- **Unit tests** (`tests/unit/`) — TypeScript tests for core functions (utilities, selection, renderers, export helpers, API client, worker-store) and MJS tests for background service worker and worker-probe summarization helpers
- **Integration tests** (`tests/integration/`) — run the export UI against a local Teams-like HTML fixture in headless Chrome via Puppeteer
- **E2E tests** — reuse the extension validation harness against a live signed-in Teams session

### Test coverage highlights

| Area | Test file | Coverage |
|------|-----------|----------|
| API message conversion | `api-client.test.ts` | Message filtering, timestamp handling, reaction conversion, author fallback |
| Conversation ID extraction | `api-client.test.ts` | URL patterns, encoded IDs, Teams Cloud root |
| Worker-store bridge | `worker-store.test.ts` | Message batches, auth tokens, members, event filtering |
| Background service worker | `background.test.mjs` | Auth headers, pagination, error responses, URL construction |
| Selection and ordering | `selection.test.ts` | Range selection, chronological ordering |
| Content extraction | `content-extraction.test.ts` | Channel time label parsing |
| Export helpers | `export-helpers.test.ts` | Labels, summaries, scope suffixes |
| Markdown rendering | `markdown-renderer.test.ts` | Quotes, reactions, mention merging |
| Utilities | `utilities.test.ts` | Text normalization, HTML escaping, color parsing |
| Full UI flow | `content-script-fixture.test.mjs` | Panel controls, selection, export, theme switching |
| Channel posts | `channel-fixture.test.mjs` | Posts, replies, subjects, mentions |
| Thread replies | `thread-replies-fixture.test.mjs` | Original post capture, reply export |

## Release

Pushing to the `main` branch triggers a GitHub Actions workflow that builds, packages, and publishes the extension zip as a GitHub Release. The version is read from `extension-src/manifest.json`.
