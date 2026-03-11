# Teams Selected Messages Export

This repository holds a browser-side Microsoft Teams message export tool for the web client, plus the local automation used to validate it against a live Teams session in Chrome.

Current scope:

- MV3 unpacked extension build
- Export controls embedded directly into the Teams top-right title bar, mounted in the header slot just left of the native settings/avatar cluster
- Collapsible export panel plus always-available quick Markdown copy button
- Opening the panel enables selection mode automatically; the panel copy action is only shown while selection mode is on, and copying Markdown turns selection mode off but preserves the selection
- Row click selection plus shift-click range selection
- Visible checkboxes with shift-click range support
- Full chat history Markdown and HTML export via an in-page upward scroll harvest that collects virtualized rows into memory before download
- Markdown export with inline `@mentions`, reply blockquotes, reactions, image placeholders, and local human-readable export timestamps
- HTML export with quoted replies, reaction chips, image placeholders, and local human-readable export timestamps
- Extension UI theme follows Teams automatically, using Teams root theme classes first and page background luminance as fallback
- Local Chrome validation harness for both injected and installed flows
- Notes from live validation against the Teams web app

Primary file:

- `src/teams-export-prototype.js`

Validation notes:

- `docs/findings.md`

## Running

### Injection mode

1. Open Microsoft Teams in Chrome.
2. Paste or inject the contents of `src/teams-export-prototype.js`.
3. Use the `Export` control injected into the Teams top bar.

### Live validation against the injected script

1. Install dependencies with `npm install`.
2. Launch a dedicated Chrome instance with `npm run chrome:launch`.
3. Sign into Teams in that Chrome window and open a chat.
4. Run `npm run validate:live`.

The validator will:

- connect to Chrome over the DevTools protocol
- inject the prototype into the live Teams page
- enable selection mode
- validate click plus shift-click range selection on the latest visible messages
- save HTML and Markdown exports to `artifacts/exports`
- save screenshots and a JSON summary under `artifacts`

The injected script is written as a self-contained IIFE and exposes a page-world debug API on `window.__teamsMessageExporter`.

### Unpacked extension flow

1. Build the extension with `npm run build:extension`.
2. Start or reuse a Chrome instance with remote debugging and Teams open.
3. Load the unpacked extension into that Chrome instance with `npm run extension:load`.
4. Run `npm run validate:extension`.

Optional:

- `npm run chrome:launch:extension` launches a dedicated Chrome profile with `extension-dist` preloaded on debug port `9223`.
- Set `DEBUG_PORT=9223` when using the dedicated extension launcher.

The installed validator uses real browser clicks against the live Teams page, verifies range selection, uses the quick-copy affordance, and downloads both Markdown and HTML exports into `artifacts/exports`.
It also runs a full-chat-history Markdown export smoke against the live Teams thread.

### Experimental worker interception probe

1. Start or reuse the Chrome debug instance with Teams already signed in.
2. Run `npm run probe:worker`.

The probe reloads Teams, patches `fetch`, `XMLHttpRequest`, `WebSocket`, and `Worker` at page start, and writes a structured summary to `artifacts/worker-intercept-probe.json`.

## Testing

- `npm run test:unit` covers pure worker-probe summarization helpers.
- `npm run test:integration` runs the export UI against a local Teams-like fixture page in Chrome/Chromium.
- `npm run test:e2e` reuses the live Teams validator and requires a signed-in Chrome debug session.
- `npm test` runs the unit and integration suite.

## Packaging

- `npm run build:extension` builds the unpacked extension into `extension-dist`.
- `npm run package:extension` builds and zips the current extension to `artifacts/teams-message-export-extension.zip`.
- The latest local package artifact is `artifacts/teams-message-export-extension.zip`.

## Current limitations

- Reactions in the shipped extension still come from DOM scraping, so actor names remain best-effort.
- Message export still depends on the visible Teams DOM and current row selectors.
- Full chat history export currently relies on DOM scrolling plus virtualized-row harvesting rather than the future worker/GraphQL data source.
- The worker/GraphQL path has been investigated and documented, but it is not wired into the extension yet.

## Status

This is now packaged as an unpacked MV3 extension for Chrome development, with the original injection harness retained for fast DOM exploration. It is not yet prepared for Chrome Web Store packaging.
