// ── API message fetching ──────────────────────────────────────────────────────
//
// The content script captures an auth token via the MAIN-world worker hook
// (reads MSAL IC3 Bearer JWT from localStorage) and passes it to the background
// when requesting a full chat export. The background makes the cross-origin API
// calls since content scripts cannot reach *.ng.msg.teams.microsoft.com directly.
//
// Two token types are supported:
//   "bearer"      → Authorization: Bearer <jwt>     (MSAL IC3 token)
//   "skypetoken"  → Authentication: skypetoken=<token>  (legacy)

const API_PAGE_SIZE = 200;
const API_MAX_PAGES = 100;

async function fetchChatMessages(token, tokenType, region, conversationId) {
  const baseUrl = `https://${region}.ng.msg.teams.microsoft.com/v1`;
  const encodedConversationId = encodeURIComponent(conversationId);
  let nextUrl = `${baseUrl}/users/ME/conversations/${encodedConversationId}/messages?pageSize=${API_PAGE_SIZE}`;
  const allMessages = [];
  let pageCount = 0;

  const authHeaders =
    tokenType === "bearer"
      ? { Authorization: `Bearer ${token}` }
      : { Authentication: `skypetoken=${token}` };

  while (nextUrl && pageCount < API_MAX_PAGES) {
    const response = await fetch(nextUrl, { headers: authHeaders });

    if (response.status === 401 || response.status === 403) {
      return {
        error: `Authentication failed (${response.status}). Token may have expired — reload Teams and try again.`
      };
    }

    if (!response.ok) {
      return {
        error: `API returned ${response.status} ${response.statusText}`
      };
    }

    const data = await response.json();

    if (Array.isArray(data.messages)) {
      allMessages.push(...data.messages);
    }

    nextUrl =
      data._metadata && data._metadata.backwardLink
        ? data._metadata.backwardLink
        : null;
    pageCount++;
  }

  return { messages: allMessages, pageCount };
}

// ── Message handler ───────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "teams-export/fetch-full-chat") {
    const { skypeToken, tokenType, region, conversationId } = message;

    if (!skypeToken || !region) {
      sendResponse({ error: "Missing API credentials (token or region)." });
      return true;
    }

    if (!conversationId) {
      sendResponse({ error: "Missing conversation ID." });
      return true;
    }

    fetchChatMessages(skypeToken, tokenType || "skypetoken", region, conversationId)
      .then(sendResponse)
      .catch((error) => sendResponse({ error: String(error) }));

    return true; // keep sendResponse channel open for async reply
  }
});

// ── Extension icon click: toggle the export panel ─────────────────────────────

chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id) {
    return;
  }

  try {
    await chrome.tabs.sendMessage(tab.id, {
      type: "teams-export/toggle-panel"
    });
  } catch (error) {
    console.warn(
      "Unable to toggle Teams Selected Messages Export panel.",
      error
    );
  }
});
