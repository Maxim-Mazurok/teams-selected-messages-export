# Repository Instructions

- Commit completed changes before every user handover.
- Use clear, scoped commit messages that describe the delivered change.

## Reverse Engineering Teams Internals

This section documents the techniques and patterns used to reverse-engineer Microsoft Teams' web client. Use these instructions when extending the extension to support new Teams features.

### Tooling

- **Playwright MCP** is the primary browser automation tool. Use it to navigate Teams, inspect DOM, execute JavaScript in the page context, and probe API endpoints.
- **Chrome DevTools Protocol (CDP)** on debug port `9223` gives raw access to network, DOM, and console. Use `http://localhost:9223/json/list` to enumerate targets (pages, service workers).
- **Chrome launch flags**: `--remote-debugging-port=9223 --load-extension=<path>/extension-dist --no-first-run --user-data-dir=/tmp/chrome-debug-profile`
- To force a clean extension reload, delete the profile entirely: `rm -rf /tmp/chrome-debug-profile` then relaunch Chrome.

### Auth Token Discovery

Teams stores MSAL tokens in `localStorage`. The extension reads them from the MAIN world via `worker-hook.js`:

1. **IC3 Bearer token**: `localStorage` key matching `*accesstoken*ic3.teams.office.com*` — parse the JSON value, extract the `secret` field. This is the Bearer token for chat/channel API calls.
2. **SKYPE-TOKEN**: `localStorage` key matching `*accesstoken*skypetoken*` — the `secret` field is a legacy `skypetoken=` credential. Also contains `regionGtms.chatService` which determines the API region (e.g., `apac`, `emea`, `amer`).
3. **Token refresh**: Tokens expire. Reloading the Teams page refreshes them. The worker-hook re-reads tokens every 5 seconds.

### API Endpoint Patterns

Base URL: `https://{region}.ng.msg.teams.microsoft.com`

| Endpoint | Description |
|----------|-------------|
| `/v1/users/ME/conversations/{conversationId}/messages?pageSize=200` | Fetch messages (DMs, group chats, channels). Paginate via `_metadata.backwardLink`. |
| `/v1/users/ME/conversations/{channelId};messageid={threadRootId}/messages?pageSize=200` | Fetch replies for a specific channel thread. |

**Important**: The endpoint `/messages/{threadId}/replies` returns 404. Always use the `;messageid=` URL parameter pattern for thread replies.

### Conversation ID Patterns

| Type | Pattern | Example |
|------|---------|---------|
| DM / Group chat | `19:{uuid}@unq.gbl.spaces` | `19:abc12345-6789-...@unq.gbl.spaces` |
| Channel (new) | `19:{hex32}@thread.tacv2` | `19:2e4b10d55f28442d92e525c12f4275c4@thread.tacv2` |
| Channel (legacy) | `19:{hex32}@thread.skype` | `19:abcdef0123456789...@thread.skype` |
| Older format | `19:{uuid}@thread.v2` | `19:abc12345-6789-...@thread.v2` |

The canonical regex: `/19:[a-f0-9-]+(?:_[a-f0-9-]+)?@(?:unq\.gbl\.spaces\|thread\.(?:v2\|tacv2\|skype))/i`

### DOM Extraction Techniques

Teams Cloud (`teams.cloud.microsoft`) uses no path-based routing — the URL stays the same regardless of which chat/channel is open. Conversation IDs must be extracted from the DOM.

**Approach priority order:**

1. **Chat title ancestor traversal** — find `[data-tid="chat-title"]`, walk up to find an ancestor with an ID containing the conversation pattern.
2. **Selected chat-list-item** — look for `[data-tid="chat-list-item"][aria-selected="true"]`, extract conversation ID from its `id` attribute.
3. **Message pane runway parents** — find `[data-tid="message-pane-list-runway"]`, walk up through ancestors looking for IDs with conversation pattern.
4. **Chat header scan** — find elements with IDs starting with `chat-header-19:`, extract the conversation ID suffix.
5. **Active channel treeitem** — for channels: `[role="treeitem"][tabindex="0"][data-testid*="channel-list-item-19:"]` — the `data-testid` contains the channel conversation ID.
6. **Channel ID fallback** — any element with `[id*="channel-list-item-19:"]`.

### Channel Message Structure

Channel API responses return all messages (root posts and replies) in a flat list. Key patterns:

- **Thread membership**: Each message has a `conversationLink` field. For channel messages, it contains `;messageid={parentThreadId}` identifying which thread the message belongs to.
- **Root posts**: Messages where the extracted thread ID equals the message's own `id` (or has no thread ID in `conversationLink`).
- **Replies**: Messages where the extracted thread ID differs from the message's own `id`.
- **Thread grouping**: Group all messages by thread ID, place root post first, sort replies chronologically within each thread, then concatenate all threads.

### Service Worker Considerations

- The extension's MV3 background service worker can be dormant or fail to start in automated Chrome sessions (Playwright, CDP). Always implement a fallback.
- **Direct fetch fallback**: Content scripts with `host_permissions` in the manifest can make cross-origin fetches directly (Chrome 95+). This bypasses the background service worker entirely.
- Teams' own service worker only intercepts requests to its own origin, not to `*.ng.msg.teams.microsoft.com`, so direct fetches from the content script work without interference.
- Use `chrome.runtime.sendMessage` with retry logic (3 attempts, exponential backoff) to attempt the background worker first, then fall back to direct fetch.

### Worker Interception (MAIN World)

Teams routes most data through a Web Worker (`/v2/worker/precompiled-web-worker-*.js`). The `worker-hook.js` script patches `window.Worker` at `document_start` to intercept:

- **GraphQL operations**: `ComponentsChatQueriesMessageListQuery`, `ComponentsChatQueriesChatQuery`, etc.
- **Structured messages**: `content`, `quotedMessages`, `fromUser.displayName`, `emotionsSummary`, `emotions[].users[].userId`
- **IndexedDB cache hits**: Worker responses come from `indexedDB_NewGetRangeMethod` when Teams serves from cache.

Cross-world communication uses `window.postMessage` with `{ source: "teams-export-worker-hook" }` as the discriminator.

### Debugging Workflow

1. Launch Chrome with debug port and extension: `npm run chrome:launch:extension`
2. Navigate to Teams, sign in, open target chat/channel.
3. Use Playwright MCP `browser_snapshot` to inspect page structure.
4. Use `browser_evaluate` to run JavaScript in the page context (e.g., read `localStorage`, query DOM).
5. Test API calls with `fetch()` in `browser_evaluate` using discovered tokens.
6. Check CDP targets at `http://localhost:9223/json/list` to verify service worker status.
7. After code changes: rebuild (`node scripts/bundle.mjs && node scripts/build-extension.mjs`), then either reload the extension or clear the Chrome profile and relaunch.

### Common Pitfalls

- **Stale content scripts**: Chrome caches injected content scripts aggressively. After rebuilding, you may need to delete the entire profile directory and relaunch Chrome to pick up changes. Verify by checking console log line numbers.
- **Trusted Types**: Teams enforces Trusted Types. Never use `innerHTML` for DOM manipulation — use `createElement` and `textContent` instead.
- **Cross-world isolation**: Properties set on `window` in the ISOLATED world are invisible to the MAIN world (and vice versa). Use `document.documentElement.dataset` or `window.postMessage` for cross-world data transfer.
- **Token expiry**: IC3 Bearer tokens expire. If API calls return 401, reload Teams to refresh the token.
- **Channel vs DM detection**: Use `@thread.tacv2` or `@thread.skype` suffix to detect channel conversations. DMs/group chats use `@unq.gbl.spaces`.
