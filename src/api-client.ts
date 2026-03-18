/**
 * API Client
 *
 * Communicates with the background service worker to:
 *   1. Check whether a skypeToken has been captured from network traffic
 *   2. Request full chat message history via the Teams Chat Service REST API
 *   3. Convert raw API messages into MessageSnapshot records for export
 *
 * The background service worker intercepts outgoing requests via chrome.webRequest
 * to passively capture the skypeToken and API region. When the content script
 * requests a full chat export, the background fetches all pages of messages from
 * the REST API and returns the raw data here for formatting.
 */

import type {
  ApiCredentials,
  ApiFetchResult,
  ApiMessage,
  MessageSnapshot,
  ReactionInfo
} from "./types.js";
import { log } from "./state.js";

// ── Chrome runtime helpers ────────────────────────────────────────────────────

function getChromeRuntime(): typeof chrome.runtime | null {
  const global = globalThis as unknown as Record<string, unknown>;
  const chromeNamespace = global.chrome as
    | { runtime?: typeof chrome.runtime }
    | undefined;
  return chromeNamespace?.runtime ?? null;
}

function sendBackgroundMessage<T>(message: Record<string, unknown>): Promise<T> {
  return new Promise((resolve, reject) => {
    const runtime = getChromeRuntime();
    if (!runtime?.sendMessage) {
      reject(new Error("chrome.runtime.sendMessage is not available"));
      return;
    }

    runtime.sendMessage(message, (response: T) => {
      if (runtime.lastError) {
        reject(new Error(runtime.lastError.message));
        return;
      }
      resolve(response);
    });
  });
}

// ── Public API ────────────────────────────────────────────────────────────────

export async function getApiCredentials(): Promise<ApiCredentials> {
  return sendBackgroundMessage<ApiCredentials>({
    type: "teams-export/get-api-credentials"
  });
}

export function isApiAvailable(): Promise<boolean> {
  return getApiCredentials()
    .then((credentials) => credentials.hasToken)
    .catch(() => false);
}

export async function fetchFullChatViaApi(
  conversationId: string
): Promise<ApiFetchResult> {
  return sendBackgroundMessage<ApiFetchResult>({
    type: "teams-export/fetch-full-chat",
    conversationId
  });
}

// ── Conversation ID extraction ────────────────────────────────────────────────

/**
 * Attempt to extract the current Teams conversation ID from the page URL.
 *
 * Teams URL patterns observed:
 *   - /conversations/{id}
 *   - /l/message/{id}/...
 *   - /l/chat/{id}/...
 *   - Hash-based: #/conversations/{id} or #/l/message/{id}
 */
export function extractConversationIdFromUrl(): string | null {
  const fullPath = `${location.pathname}${location.hash}`;

  const patterns = [
    /\/conversations\/([^/?&#]+)/,
    /\/l\/(?:message|chat)\/([^/?&#]+)/
  ];

  for (const pattern of patterns) {
    const match = fullPath.match(pattern);
    if (match) {
      return decodeURIComponent(match[1]);
    }
  }

  return null;
}

/**
 * Get the conversation ID from the URL or from the background's tracked API
 * requests.
 */
export async function resolveConversationId(): Promise<string | null> {
  const fromUrl = extractConversationIdFromUrl();
  if (fromUrl) {
    return fromUrl;
  }

  const credentials = await getApiCredentials().catch(() => null);
  return credentials?.lastApiConversationId ?? null;
}

// ── API message → MessageSnapshot conversion ─────────────────────────────────

/** Filter out system/control messages that shouldn't appear in exports. */
const EXCLUDED_MESSAGE_TYPES = new Set([
  "ThreadActivity/AddMember",
  "ThreadActivity/MemberJoined",
  "ThreadActivity/MemberLeft",
  "ThreadActivity/DeleteMember",
  "ThreadActivity/TopicUpdate",
  "ThreadActivity/HistoryDisclosedUpdate",
  "Event/Call",
  "RichText/Media_CallRecording",
  "MessageDelete"
]);

function stripHtmlTags(html: string): string {
  const temporary = document.createElement("div");
  temporary.innerHTML = html;
  return temporary.textContent?.trim() ?? "";
}

function formatApiTimestamp(isoTimestamp: string | undefined): {
  label: string;
  dateTime: string;
} {
  if (!isoTimestamp) {
    return { label: "", dateTime: "" };
  }

  const date = new Date(isoTimestamp);
  if (Number.isNaN(date.getTime())) {
    return { label: isoTimestamp, dateTime: isoTimestamp };
  }

  const label = date.toLocaleString(undefined, {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });

  return { label, dateTime: date.toISOString() };
}

function convertApiReactions(emotions: ApiMessage["emotions"]): ReactionInfo[] {
  if (!emotions || !Array.isArray(emotions)) {
    return [];
  }

  return emotions
    .filter((emotion) => emotion.users && emotion.users.length > 0)
    .map((emotion) => ({
      emoji: emotion.key,
      name: emotion.key,
      count: emotion.users.length,
      actors: emotion.users
        .map((user) => user.displayName ?? user.mri)
        .filter(Boolean)
    }));
}

/**
 * Convert a plain-text content body to basic Markdown. Handles Teams' HTML
 * message format (RichText/Html) by extracting text, and passes through plain
 * text messages directly.
 */
function convertApiContentToMarkdown(
  content: string | undefined,
  messageType: string | undefined
): string {
  if (!content) {
    return "";
  }

  if (messageType === "RichText/Html" || content.includes("<")) {
    return stripHtmlTags(content);
  }

  return content;
}

export function convertApiMessagesToSnapshots(
  apiMessages: ApiMessage[]
): MessageSnapshot[] {
  const snapshots: MessageSnapshot[] = [];

  const filteredMessages = apiMessages.filter(
    (message) => !EXCLUDED_MESSAGE_TYPES.has(message.messagetype ?? message.type ?? "")
  );

  // API returns newest-first per page; reverse for chronological order
  const chronological = [...filteredMessages].reverse();

  for (let index = 0; index < chronological.length; index++) {
    const apiMessage = chronological[index];
    const timestamp = formatApiTimestamp(
      apiMessage.originalarrivaltime ?? apiMessage.composetime
    );
    const content = apiMessage.content ?? "";
    const messageType = apiMessage.messagetype ?? apiMessage.type;

    snapshots.push({
      id: apiMessage.id,
      index,
      author: apiMessage.imdisplayname ?? "Unknown author",
      timeLabel: timestamp.label,
      dateTime: timestamp.dateTime,
      subject: (apiMessage.properties?.subject as string) ?? "",
      quote: null,
      reactions: convertApiReactions(apiMessage.emotions),
      html: messageType === "RichText/Html" ? content : `<p>${content}</p>`,
      markdown: convertApiContentToMarkdown(content, messageType),
      plainText: stripHtmlTags(content),
      captureOrder: index
    });
  }

  return snapshots;
}

// ── High-level fetch + convert ────────────────────────────────────────────────

export async function fetchAndConvertFullChat(
  conversationId: string,
  onProgress?: (text: string) => void
): Promise<{ snapshots: MessageSnapshot[]; pageCount: number } | { error: string }> {
  onProgress?.("Connecting to Teams API...");

  const result = await fetchFullChatViaApi(conversationId);

  if (result.error) {
    log(`[api-client] API fetch failed: ${result.error}`);
    return { error: result.error };
  }

  if (!result.messages || result.messages.length === 0) {
    return { error: "API returned no messages." };
  }

  onProgress?.(
    `Fetched ${result.messages.length} messages in ${result.pageCount ?? "?"} pages, converting...`
  );

  const snapshots = convertApiMessagesToSnapshots(result.messages);

  log(
    `[api-client] Converted ${snapshots.length} messages from ${result.messages.length} raw API messages`
  );

  return { snapshots, pageCount: result.pageCount ?? 0 };
}
