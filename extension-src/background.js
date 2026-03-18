// ── Token capture state ────────────────────────────────────────────────────────

let capturedSkypeToken = null;
let capturedRegion = null;

// ── Intercept requests to capture the skype token ─────────────────────────────

chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    const urlMatch = details.url.match(
      /https:\/\/(\w+)\.ng\.msg\.teams\.microsoft\.com/
    );
    if (urlMatch) {
      capturedRegion = urlMatch[1];
    }

    if (!details.requestHeaders) {
      return;
    }

    for (const header of details.requestHeaders) {
      const name = header.name.toLowerCase();

      if (name === "authentication" && header.value) {
        const prefix = "skypetoken=";
        if (header.value.startsWith(prefix)) {
          capturedSkypeToken = header.value.slice(prefix.length);
        }
      }

      // Fallback: some requests use x-skypetoken header directly
      if (name === "x-skypetoken" && header.value && !capturedSkypeToken) {
        capturedSkypeToken = header.value;
      }
    }
  },
  { urls: ["https://*.ng.msg.teams.microsoft.com/*"] },
  ["requestHeaders", "extraHeaders"]
);

// Also capture from teams.cloud.microsoft requests (x-skypetoken header)
chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    if (!details.requestHeaders) {
      return;
    }

    for (const header of details.requestHeaders) {
      if (header.name.toLowerCase() === "x-skypetoken" && header.value) {
        capturedSkypeToken = header.value;

        // Extract region from URL pattern like /api/mt/apac/...
        const regionMatch = details.url.match(/\/api\/(?:mt|chatsvc)\/(\w+)\//);
        if (regionMatch && !capturedRegion) {
          capturedRegion = regionMatch[1];
        }
      }
    }
  },
  {
    urls: [
      "https://teams.cloud.microsoft/*",
      "https://teams.microsoft.com/*"
    ]
  },
  ["requestHeaders", "extraHeaders"]
);

// ── Conversation ID tracking from intercepted API URLs ────────────────────────

let lastApiConversationId = null;

chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    const conversationMatch = details.url.match(
      /\/conversations\/([^/?]+)\/messages/
    );
    if (conversationMatch) {
      lastApiConversationId = decodeURIComponent(conversationMatch[1]);
    }
  },
  { urls: ["https://*.ng.msg.teams.microsoft.com/*"] },
  []
);

// ── API message fetching ──────────────────────────────────────────────────────

const API_PAGE_SIZE = 200;
const API_MAX_PAGES = 100;

async function fetchChatMessages(conversationId) {
  if (!capturedSkypeToken || !capturedRegion) {
    return {
      error: "No API credentials captured yet. Navigate to a Teams chat first."
    };
  }

  const baseUrl = `https://${capturedRegion}.ng.msg.teams.microsoft.com/v1`;
  const encodedConversationId = encodeURIComponent(conversationId);
  let nextUrl = `${baseUrl}/users/ME/conversations/${encodedConversationId}/messages?pageSize=${API_PAGE_SIZE}`;
  const allMessages = [];
  let pageCount = 0;

  while (nextUrl && pageCount < API_MAX_PAGES) {
    const response = await fetch(nextUrl, {
      headers: {
        Authentication: `skypetoken=${capturedSkypeToken}`
      }
    });

    if (response.status === 401 || response.status === 403) {
      capturedSkypeToken = null;
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
  if (message.type === "teams-export/toggle-panel") {
    if (!sender.tab || !sender.tab.id) {
      return;
    }
    chrome.tabs.sendMessage(sender.tab.id, {
      type: "teams-export/toggle-panel"
    });
    return;
  }

  if (message.type === "teams-export/get-api-credentials") {
    sendResponse({
      hasToken: Boolean(capturedSkypeToken),
      region: capturedRegion,
      lastApiConversationId
    });
    return true;
  }

  if (message.type === "teams-export/fetch-full-chat") {
    const conversationId =
      message.conversationId || lastApiConversationId;

    if (!conversationId) {
      sendResponse({
        error: "Could not determine conversation ID. Open a chat and try again."
      });
      return true;
    }

    fetchChatMessages(conversationId)
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
