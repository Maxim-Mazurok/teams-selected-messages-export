/**
 * Direct Teams API Client
 *
 * Captures a Skype auth token from a running Teams session (via CDP Fetch
 * interception) and makes direct HTTP calls to the Teams Chat Service REST API.
 *
 * Flow:
 *   1. Connect to a Chrome debug session with Teams open
 *   2. Reload the page to trigger token acquisition
 *   3. Capture the x-skypetoken from an intercepted request header
 *   4. Use the token to call {region}.ng.msg.teams.microsoft.com/v1/...
 *
 * Key discovery:
 *   - Teams stores auth tokens encrypted in localStorage; not directly extractable
 *   - The refresh token is SPA-bound (AADSTS9002327) and cannot be redeemed from Node.js
 *   - The only reliable extraction path is intercepting live request headers via CDP Fetch
 *   - The Skype token alone (Authentication: skypetoken=<token>) is sufficient for
 *     all chat service calls: conversations, messages, members
 *   - Token lifetime is ~24 hours based on observed expiry timestamps
 */

import type { Page, CDPSession, Protocol } from "puppeteer-core";

export interface TeamsAuthToken {
  skypeToken: string;
  region: string;
}

export interface TeamsConversation {
  id: string;
  topic: string;
  threadType: string;
  version: number;
  lastMessageTime: string | null;
  memberCount: number | null;
}

export interface TeamsApiMessage {
  id: string;
  messageType: string;
  from: string;
  displayName: string;
  content: string;
  originalArrivalTime: string;
  composeTime: string;
  editTime: string | null;
  subject: string | null;
  isDeleted: boolean;
  emotions: TeamsApiEmotion[];
  mentions: TeamsApiMention[];
  quotedMessageId: string | null;
}

export interface TeamsApiEmotion {
  key: string;
  users: Array<{ mri: string; time: number }>;
}

export interface TeamsApiMention {
  id: string;
  displayName: string;
}

export interface TeamsApiMember {
  id: string;
  displayName: string;
  role: string;
}

export interface TeamsMessagesPage {
  messages: TeamsApiMessage[];
  backwardLink: string | null;
  syncState: string | null;
}

const CHAT_SERVICE_BASE = (region: string) =>
  `https://${region}.ng.msg.teams.microsoft.com/v1`;

const FETCH_INTERCEPT_TIMEOUT = 20_000;
const PAGE_RELOAD_TIMEOUT = 25_000;

/**
 * Capture a Skype token from a running Teams page by intercepting
 * request headers via the CDP Fetch domain during a page reload.
 */
export async function captureSkypeToken(
  teamsPage: Page,
): Promise<TeamsAuthToken> {
  const cdpSession: CDPSession = await teamsPage.createCDPSession();
  let skypeToken: string | null = null;

  await cdpSession.send("Fetch.enable", {
    patterns: [{ urlPattern: "*teams*", requestStage: "Request" }],
  });

  const tokenPromise = new Promise<void>((resolve) => {
    const timeout = setTimeout(resolve, FETCH_INTERCEPT_TIMEOUT);

    cdpSession.on("Fetch.requestPaused", async (event: Protocol.Fetch.RequestPausedEvent) => {
      const headers = event.request.headers ?? {};
      const requestId = event.requestId;

      for (const [name, value] of Object.entries(headers)) {
        if (name.toLowerCase() === "x-skypetoken" && !skypeToken) {
          skypeToken = value;
        }
      }

      try {
        await cdpSession.send("Fetch.continueRequest", { requestId });
      } catch {
        // Request may have already been handled
      }

      if (skypeToken) {
        clearTimeout(timeout);
        resolve();
      }
    });
  });

  // Reload the page to trigger Teams auth flow
  teamsPage
    .reload({ waitUntil: "networkidle2", timeout: PAGE_RELOAD_TIMEOUT })
    .catch(() => {
      // Reload may time out, but we should have the token by then
    });

  await tokenPromise;

  await cdpSession.send("Fetch.disable");
  await cdpSession.detach();

  if (!skypeToken) {
    throw new Error(
      "Failed to capture Skype token. Ensure Teams is loaded and authenticated.",
    );
  }

  // Detect region from the Teams page URL or default to apac
  const region = detectRegion(teamsPage.url());

  return { skypeToken, region };
}

/**
 * Find the Teams page in a connected browser.
 */
export function findTeamsPage(pages: Page[]): Page | undefined {
  return pages.find((page) =>
    /teams\.(microsoft|cloud\.microsoft)/i.test(page.url()),
  );
}

/**
 * Detect the API region from Teams configuration.
 * Falls back to "apac" if not determinable.
 */
function detectRegion(_pageUrl: string): string {
  // In production, we could detect from apac/emea/amer in API URLs
  // or from the user's timezone. For now, all regions work identically
  // (they all return the same data), so we default to apac.
  return "apac";
}

function authHeaders(token: TeamsAuthToken): Record<string, string> {
  return {
    Authentication: `skypetoken=${token.skypeToken}`,
  };
}

/**
 * Fetch the user's conversation list.
 */
export async function fetchConversations(
  token: TeamsAuthToken,
  pageSize = 50,
): Promise<TeamsConversation[]> {
  const url = `${CHAT_SERVICE_BASE(token.region)}/users/ME/conversations?view=mychats&pageSize=${pageSize}`;

  const response = await fetch(url, { headers: authHeaders(token) });
  if (!response.ok) {
    throw new Error(`Failed to fetch conversations: ${response.status} ${response.statusText}`);
  }

  const data = (await response.json()) as {
    conversations: Array<{
      id: string;
      version: number;
      threadProperties?: { topic?: string; threadType?: string; memberCount?: string };
      properties?: { displayName?: string; lastimreceivedtime?: string };
    }>;
  };

  return (data.conversations ?? []).map((conversation) => ({
    id: conversation.id,
    topic:
      conversation.threadProperties?.topic ??
      conversation.properties?.displayName ??
      "",
    threadType: conversation.threadProperties?.threadType ?? "chat",
    version: conversation.version,
    lastMessageTime: conversation.properties?.lastimreceivedtime ?? null,
    memberCount: conversation.threadProperties?.memberCount
      ? Number(conversation.threadProperties.memberCount)
      : null,
  }));
}

/**
 * Fetch messages from a specific conversation.
 * Returns one page of messages plus a backward link for pagination.
 */
export async function fetchMessages(
  token: TeamsAuthToken,
  conversationId: string,
  pageSize = 50,
  backwardLink?: string,
): Promise<TeamsMessagesPage> {
  const url =
    backwardLink ??
    `${CHAT_SERVICE_BASE(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages?pageSize=${pageSize}`;

  const response = await fetch(url, { headers: authHeaders(token) });
  if (!response.ok) {
    throw new Error(`Failed to fetch messages: ${response.status} ${response.statusText}`);
  }

  const data = (await response.json()) as {
    messages: Array<Record<string, unknown>>;
    _metadata?: { backwardLink?: string; syncState?: string };
  };

  const messages = (data.messages ?? []).map(parseApiMessage);

  return {
    messages,
    backwardLink: data._metadata?.backwardLink ?? null,
    syncState: data._metadata?.syncState ?? null,
  };
}

/**
 * Fetch all messages from a conversation by following pagination links.
 */
export async function fetchAllMessages(
  token: TeamsAuthToken,
  conversationId: string,
  options: { maxPages?: number; pageSize?: number; onProgress?: (count: number) => void } = {},
): Promise<TeamsApiMessage[]> {
  const maxPages = options.maxPages ?? 100;
  const pageSize = options.pageSize ?? 200;
  const allMessages: TeamsApiMessage[] = [];
  let backwardLink: string | undefined;

  for (let page = 0; page < maxPages; page++) {
    const result = await fetchMessages(token, conversationId, pageSize, backwardLink);
    allMessages.push(...result.messages);

    options.onProgress?.(allMessages.length);

    if (!result.backwardLink) break;
    backwardLink = result.backwardLink;
  }

  return allMessages;
}

/**
 * Fetch members of a conversation.
 */
export async function fetchMembers(
  token: TeamsAuthToken,
  conversationId: string,
): Promise<TeamsApiMember[]> {
  const url = `${CHAT_SERVICE_BASE(token.region)}/threads/${encodeURIComponent(conversationId)}/members`;

  const response = await fetch(url, { headers: authHeaders(token) });
  if (!response.ok) {
    throw new Error(`Failed to fetch members: ${response.status} ${response.statusText}`);
  }

  const data = (await response.json()) as {
    members: Array<{
      id: string;
      userDisplayName?: string;
      role?: string;
    }>;
  };

  return (data.members ?? []).map((member) => ({
    id: member.id,
    displayName: member.userDisplayName ?? "",
    role: member.role ?? "member",
  }));
}

/**
 * Parse a raw API message into our structured format.
 */
function parseApiMessage(raw: Record<string, unknown>): TeamsApiMessage {
  const properties = (raw.properties ?? {}) as Record<string, unknown>;

  // Parse emotions from properties
  const rawEmotions = properties.emotions as
    | Array<{ key: string; users: Array<{ mri: string; time: number }> }>
    | string
    | undefined;

  let emotions: TeamsApiEmotion[] = [];
  if (typeof rawEmotions === "string") {
    try {
      emotions = JSON.parse(rawEmotions) as TeamsApiEmotion[];
    } catch {
      emotions = [];
    }
  } else if (Array.isArray(rawEmotions)) {
    emotions = rawEmotions;
  }

  // Parse mentions
  const rawMentions = properties.mentions as
    | Array<{ id: string; displayName?: string }>
    | string
    | undefined;

  let mentions: TeamsApiMention[] = [];
  if (typeof rawMentions === "string") {
    try {
      const parsed = JSON.parse(rawMentions) as Array<{ id: string; displayName?: string }>;
      mentions = parsed.map((mention) => ({
        id: mention.id ?? "",
        displayName: mention.displayName ?? "",
      }));
    } catch {
      mentions = [];
    }
  } else if (Array.isArray(rawMentions)) {
    mentions = rawMentions.map((mention) => ({
      id: mention.id ?? "",
      displayName: mention.displayName ?? "",
    }));
  }

  // Detect quoted messages from content (reply format uses skype schema)
  let quotedMessageId: string | null = null;
  const content = String(raw.content ?? "");
  const quoteMatch = content.match(
    /itemtype="http:\/\/schema\.skype\.com\/Reply"\s+itemid="(\d+)"/,
  );
  if (quoteMatch) {
    quotedMessageId = quoteMatch[1];
  }

  return {
    id: String(raw.id ?? ""),
    messageType: String(raw.messagetype ?? ""),
    from: String(raw.from ?? ""),
    displayName: String(raw.imdisplayname ?? ""),
    content,
    originalArrivalTime: String(raw.originalarrivaltime ?? ""),
    composeTime: String(raw.composetime ?? ""),
    editTime: raw.properties
      ? String((properties.edittime as string) ?? "")
      : null,
    subject: properties.subject ? String(properties.subject) : null,
    isDeleted: raw.messagetype === "MessageDelete" || raw.properties
      ? String(properties.deletetime ?? "") !== ""
      : false,
    emotions,
    mentions,
    quotedMessageId,
  };
}

/**
 * Send a message to a conversation.
 * Returns the server-assigned message ID.
 */
export async function sendMessage(
  token: TeamsAuthToken,
  conversationId: string,
  messageContent: string,
  senderDisplayName: string,
): Promise<string> {
  const url = `${CHAT_SERVICE_BASE(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`;

  const clientMessageId = String(Date.now());

  const body = {
    content: messageContent,
    messagetype: "Text",
    contenttype: "text",
    clientmessageid: clientMessageId,
    imdisplayname: senderDisplayName,
    properties: {
      importance: "",
      subject: null,
    },
  };

  const response = await fetch(url, {
    method: "POST",
    headers: {
      ...authHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to send message: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const data = (await response.json()) as { OriginalArrivalTime: number; id?: string };
  return data.id ?? clientMessageId;
}

/**
 * Find an existing 1:1 conversation with a specific user by scanning
 * recent messages for display names.
 *
 * The members API doesn't return display names for 1:1 chats, so instead
 * we fetch a few recent messages from each untitled chat and match against
 * the sender display names (imdisplayname).
 *
 * Also handles the self-chat case (48:notes) when the target matches
 * the logged-in user.
 *
 * Only searches 1:1 chats — use the chatFilter in the CLI script for
 * matching group chats and meetings by topic.
 */
export async function findOneOnOneConversation(
  token: TeamsAuthToken,
  targetIdentifier: string,
): Promise<{ conversationId: string; memberDisplayName: string } | null> {
  const conversations = await fetchConversations(token, 100);

  const targetLower = targetIdentifier.toLowerCase();

  // Check self-chat first — matches when the user searches for their own name
  const selfChat = conversations.find((conversation) =>
    conversation.id.startsWith("48:notes"),
  );
  if (selfChat) {
    const currentUserName = await fetchCurrentUserDisplayName(token);
    if (currentUserName.toLowerCase().includes(targetLower)) {
      return { conversationId: selfChat.id, memberDisplayName: `${currentUserName} (self)` };
    }
  }

  // Search untitled 1:1 chats by scanning recent messages for sender names
  const untitledChats = conversations.filter(
    (conversation) =>
      conversation.threadType === "chat" && !conversation.topic,
  );

  for (const chat of untitledChats) {
    try {
      const messagesPage = await fetchMessages(token, chat.id, 10);
      const textMessages = messagesPage.messages.filter(
        (message) =>
          message.messageType === "RichText/Html" ||
          message.messageType === "Text",
      );

      // Collect unique sender names
      const senderNames = [
        ...new Set(
          textMessages
            .map((message) => message.displayName)
            .filter((name) => name.length > 0),
        ),
      ];

      // Check if any sender matches the target
      for (const senderName of senderNames) {
        if (senderName.toLowerCase().includes(targetLower)) {
          return { conversationId: chat.id, memberDisplayName: senderName };
        }
      }
    } catch {
      // Some older chats may not be fetchable
      continue;
    }
  }

  return null;
}

/**
 * Get the display name of the current user from their own message history.
 *
 * Fetches recent messages from a conversation and identifies the current
 * user's display name by finding the most common sender (likely "me").
 */
export async function fetchCurrentUserDisplayName(
  token: TeamsAuthToken,
): Promise<string> {
  // The ME/properties endpoint has user settings but not always displayname.
  // The most reliable approach is to look at our own sent messages.
  const conversations = await fetchConversations(token, 10);

  for (const conversation of conversations) {
    try {
      const messagesPage = await fetchMessages(token, conversation.id, 10);
      const textMessages = messagesPage.messages.filter(
        (message) =>
          (message.messageType === "RichText/Html" ||
            message.messageType === "Text") &&
          message.displayName.length > 0,
      );

      if (textMessages.length === 0) continue;

      // Count sender occurrences — the current user likely appears most often
      // across multiple conversations. But for a single conversation, just
      // return the first sender name we find (we'll validate it's "us" later).
      // For now, use a heuristic: check the self-chat (48:notes) messages
      if (conversation.id.startsWith("48:notes")) {
        const selfMessage = textMessages[0];
        if (selfMessage) return selfMessage.displayName;
      }
    } catch {
      continue;
    }
  }

  // Fallback: try the properties endpoint
  const url = `${CHAT_SERVICE_BASE(token.region)}/users/ME/properties`;
  try {
    const response = await fetch(url, { headers: authHeaders(token) });
    if (response.ok) {
      const data = (await response.json()) as Record<string, unknown>;
      if (typeof data.displayname === "string") {
        return data.displayname;
      }
    }
  } catch {
    // Fall through to default
  }

  return "Unknown User";
}
