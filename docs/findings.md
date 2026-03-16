# Findings

This file captures live DOM notes, validation status, and extension-packaging considerations discovered while testing the Teams export tool against Microsoft Teams in Chrome.

## Validation Log

- Teams web currently exposes message rows as `div[data-tid="chat-pane-item"]`.
- Useful child selectors on message rows:
  - `data-testid="message-wrapper"`
  - `data-tid="chat-pane-message"`
  - `data-tid="message-author-name"`
  - `time[datetime]`
- `data-tid="chat-pane-item"` is reused for some non-message nodes, so filtering on `message-wrapper` plus `chat-pane-message` is required.
- Teams enforces Trusted Types on this page. Prototype code cannot rely on `innerHTML` to build UI elements inside the live page.
- `addScriptTag`-style DOM injection is not reliable here because Trusted Types blocks script text assignment.
- A local Chrome DevTools Protocol harness was added as fallback after the Playwright MCP transport became unusable during browser reset.
- Live validation succeeded on Wednesday, March 11, 2026 in a dedicated Chrome instance using the local DevTools Protocol harness.
- Validated conversation types:
  - Direct chat (1:1)
  - Group chat
- For both validations:
  - selection mode loaded successfully
  - row click selection worked
  - shift-click range selection selected 3 contiguous messages
  - Markdown export was written to `artifacts/exports`
  - HTML export was written to `artifacts/exports`
  - screenshots were captured under `artifacts/screenshots`
- Validation summary is stored at `artifacts/validation-summary.json`.
- `artifacts/validation-summary.json` is the canonical final result set. Additional export files may exist from intermediate exploratory passes.
- Reply export validation succeeded on Wednesday, March 11, 2026 against a replied message in a group chat.
- Quote-specific validation artifacts:
  - `artifacts/exports/<conversation>-quote-check-<timestamp>.md`
  - `artifacts/screenshots/<conversation>-quote-check-<timestamp>.png`
- The Markdown export now emits replied-to messages as a real blockquote with reply author and timestamp before the reply body.
- Inline mention handling was corrected so tagged users stay in-flow as `@Name` instead of being broken across lines in Markdown.
- Export timestamps now use the browser's local time zone in a human-readable long date format instead of raw ISO timestamps.
- Selection mode suppresses native browser text selection inside the Teams message runway.
- Generic inline images are exported as `[Image omitted]` placeholders instead of being ignored silently or downloaded.
- Reaction summaries are exported in both Markdown and HTML. Emoji, reaction name, and count are available reliably from the visible DOM.
- Reaction actor names are still a best-effort field only. They were not exposed consistently in the live Teams DOM during validation.
- The launcher is now always visible on the Teams page, with a collapsible panel and a quick-copy Markdown button that appears when messages are selected.
- The export controls are now embedded directly into the Teams title bar by mounting the dock into `#titlebar-end-slot-start`, which places them immediately to the left of the native `Settings and more` and profile controls instead of overlaying them.
- A fixed `top: 10px; right: 200px` fallback still exists if the Teams header slot is unavailable.
- The primary header button now stays labeled `Export`; the selected-count badge only appears on the quick-copy action (`Copy N`) to avoid redundant counts.
- The visible selector text and manual rescan control were removed from the UI. The list now auto-refreshes from DOM mutations.
- The current panel flow is:
  - opening the launcher opens the panel and enables selection mode automatically
  - the panel-level `Copy MD` action now stays visible when selection mode is off, but it is disabled until selection mode is back on with at least one selected message
  - the panel actions now sit in a uniform 3x2 grid in this order: `Copy MD | Clear Selection`, `Download MD | Download HTML`, `Full chat MD | Full chat HTML`
  - copying Markdown stops selection mode, closes the panel, and preserves the selected message set
  - reopening the panel restores the selected messages when selection mode turns back on
- The extension UI now follows Teams theme state. It reads Teams root classes such as `theme-defaultV2` first, falls back to page background luminance when explicit theme classes are absent, and updates live when the root class changes.
- Startup stabilization now re-checks both dock placement and message scanning over the first ~1.4 seconds after injection, which fixes the header-placement race that previously left the launcher floating until the first interaction.
- Selected messages now use both the left accent rail and a tinted background fill, so the selected state remains obvious even when the thread is visually dense.
- When the active Teams thread changes, the extension now resets the old selection state and rebinds against the new chat so selection can continue immediately instead of getting stuck on the previous conversation.
- The panel now exposes both `Full chat MD` and `Full chat HTML`, which scroll upward through the current conversation, accumulate message snapshots in memory, restore the approximate original viewport, and then download the harvested history in the requested format.
- Installed extension validation succeeded on Wednesday, March 11, 2026 against a group chat in Chrome on debug port `9222`.
- Installed extension artifacts are stored at:
  - `artifacts/installed-extension-validation.json`
  - `artifacts/exports/maxim-mazurok-you-installed-extension-2026-03-11T09-57-50-716Z-downloads`
  - `artifacts/screenshots/maxim-mazurok-you-installed-extension-2026-03-11T09-57-50-716Z.png`
- The latest installed validation summary now includes `panelClosedAfterPanelCopy: true` and `selectionStoppedAfterPanelCopy: true`.
- The latest installed validation summary also includes `fullChatHeadingCount: 9` and `fullChatScopePresent: true`, confirming the real `Full Chat MD` path produced a larger Markdown export than the 3-message selection smoke.
- A publishable zip for the current build is now written to `artifacts/teams-selected-messages-export-extension.zip`.
- The installed validator now uses real pointer clicks on the page instead of the page-world prototype API, which matters because extension content scripts run in an isolated world.
- A live network capture on Wednesday, March 11, 2026 showed that Teams `chatsvc` message responses already include structured reaction data in the response body:
  - `annotationsSummary.emotions`
  - `properties.emotions[].key`
  - `properties.emotions[].users[].mri`
- This means reaction actor identities are available on the wire, but not from the visible message DOM. Extracting them in an extension would likely require response-body capture through `chrome.debugger` or page-level fetch/XHR instrumentation rather than plain DOM scraping.
- A separate front-end-only probe on Wednesday, March 11, 2026 showed that patching main-thread `fetch` and `XMLHttpRequest` at page start does not capture the interesting Teams chat payloads in this app shell. The captured main-thread traffic was limited to config/bootstrap requests.
- The same probe showed that patching `window.Worker` at page start is viable. Teams creates a precompiled worker (`/v2/worker/precompiled-web-worker-*.js`) and sends GraphQL operations through `worker.postMessage`, with structured results returned on `message` events.
- The worker traffic is significantly richer than the DOM and is enough to power export features without CDP:
  - worker requests expose GraphQL operation names such as `ComponentsChatQueriesMessageListQuery`, `ComponentsChatQueriesMissingMessagesQuery`, and `DataResolversBrowserChatServiceEventsChatServiceBatchEventSubscription`
  - worker responses expose structured message objects with `content`, `quotedMessages`, `fromUser.displayName`, `emotionsSummary`, `emotions[].users[].userId`, and `diverseEmotions[].users[]`
  - in the sampled run, the message list response came from `indexedDB_NewGetRangeMethod`, which means the worker hook still sees structured data when Teams serves the thread from cache instead of the network
- The latest reproducible probe run wrote `artifacts/worker-intercept-probe.json` and observed `worker-create: 3`, `worker-request: 57`, `worker-response: 47`, with `fetch: 0`, `xhr: 0`, and `ws: 0`.
- The current extension manifest injects the UI content script at `document_idle`, which is too late for constructor patching. A worker-based implementation will need a separate `document_start` main-world hook plus a bridge back to the existing extension UI.
- The experimental probe is saved in the repo at `scripts/probe-worker-intercepts.mjs`, and the latest run writes `artifacts/worker-intercept-probe.json`.
- The experimental probe is saved in the repo at `scripts/probe-worker-intercepts.mjs`, and the latest run writes `artifacts/worker-intercept-probe.json`.
- A worker-intercept POC was built and validated on Monday, March 16, 2026 on the `feature/worker-intercept-poc` branch. All pipeline stages passed:
  - `extension-src/worker-hook.js` runs at `document_start` in `world: "MAIN"` (Chrome MV3 content script, no CDP needed).
  - `window.Worker` is patched before Teams creates its precompiled worker; `window.Worker.name` becomes `"WorkerWrapper"` confirming the patch is in place.
  - Worker response messages are intercepted, stripped to a safe summary, and forwarded to the isolated content script via `window.postMessage({ source: "teams-export-worker-hook", ... }, "*")`.
  - The isolated-world `worker-store.ts` listens for bridge events on `window.addEventListener("message", ...)` and stores messages in a `Map<messageId, WorkerCapturedMessage>`.
  - In the validated run: 6 bridge events were received in the MAIN world; 40 structured chat messages were stored in the isolated-world worker store; 9 unique member identities were accumulated from `fromUser.displayName` fields in the responses.
  - The reaction-enrichment path in `messages.ts` now calls `enrichReactionsFromWorker(messageId, domReactions)` which looks up the worker store by message ID, matches reaction keys by normalised name (e.g. DOM `"like"` ↔ worker key `"like"`), and replaces the `actors` array with resolved display names when available.
  - Message ID normalisation handles the `content-<id>` and `timestamp-<id>` DOM prefixes that Teams uses on child elements, so the join between DOM message IDs and worker message IDs is stable.
  - The `open -na "Google Chrome"` launcher script reuses an existing Chrome process and ignores extra flags; the extension must be loaded separately via `scripts/load-unpacked-extension.mjs` or by launching Chrome directly with its binary using `--load-extension`.
  - Chrome MV3 content script `world: "MAIN"` injection requires Chrome 111 or later; confirmed working on Chrome 145.
  - Properties set on `window` by content scripts (ISOLATED world) are NOT visible to the page's main world. Use DOM attributes (e.g. `document.documentElement.dataset.*`) or `window.postMessage` for cross-world communication. The extension public API (`window.__teamsMessageExporter`) is therefore not accessible from puppeteer `page.evaluate()`.
  - The worker hook POC validation is stored at `artifacts/worker-poc-validation.json` and is reproducible with `npm run validate:worker-poc`.

## Caveats

- Message row filtering must be tied to the chat runway container (`[data-tid="message-pane-list-runway"]` or `#chat-pane-list`) rather than semantic `main` landmarks.
- Sidebar chat items were present as level-2 `treeitem` nodes, but useful `aria-label` values were not available in the live Chrome DOM. Alternate chat switching uses visible item order instead.
- The selection treatment is now a soft highlight plus left accent rail instead of a full outline, which avoids the stacked-border collisions seen in the earlier prototype.
- Reliable range selection in the installed extension path depends on handling row selection on `mousedown`, because Teams and Chrome can interfere with the later `click` event on some message structures.
- Checkbox range selection is now driven from the checkbox click itself so shift-click works there as well.
- Full chat history export currently relies on DOM scrolling and snapshotting rather than the worker/GraphQL path. It works against the live validated thread, but the worker path should become the primary source of truth later because it is less sensitive to virtualization behavior.

## Worker Hook Next Steps

1. ~~Add a dedicated `document_start` main-world hook script to the extension.~~ **Done** (`extension-src/worker-hook.js`, `world: "MAIN"`).
2. ~~Patch `window.Worker` first.~~ **Done** — Worker constructor is patched; confirmed with `window.Worker.name === "WorkerWrapper"`.
3. ~~Bridge only summarized payloads back to the extension UI layer with `window.postMessage`.~~ **Done** (`worker-store.ts` listens and stores messages end-to-end).
4. Build a more complete ID map from worker responses. The current member map is populated only from message `fromUser` fields. Capturing explicit member-list responses (e.g. `ComponentsChatQueriesChatQuery`) and the `singleMe` identity would fill remaining gaps in reaction actor resolution.
5. Join worker-derived `content` fields to exports so the HTML export uses raw worker HTML instead of DOM-scraped HTML, which would be more reliable against DOM virtualization.
6. Extend validation with mixed-source checks: DOM-selected messages combined with worker-enriched reactions; test with reactions that have more than one actor to confirm display-name resolution; test against channel threads in addition to DM and group chats.
7. Consider making the worker store the primary source for full-chat export history instead of the current DOM-scroll approach, since the worker data is already cached from indexedDB and does not require virtualization-safe scrolling.

## Expected follow-up

- Lock message row selectors to the current Teams web DOM
- Confirm compatibility across 1:1 chats, group chats, and channel threads
- Resolve richer reaction actor attribution through the worker path
- Add release-ready packaging details such as store metadata and publish checklist
