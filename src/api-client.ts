/**
 * API Client
 *
 * Communicates with the background service worker to fetch full chat history
 * via the Teams Chat Service REST API.
 *
 * Token capture flow:
 *   1. worker-hook.js (MAIN world) intercepts fetch/XHR to capture x-skypetoken
 *   2. worker-store.ts receives the token via window.postMessage bridge
 *   3. This module reads the token from worker-store and passes it to background
 *   4. Background service worker makes the actual cross-origin API calls
 */

import type {
  ApiFetchResult,
  ApiMessage,
  MessageSnapshot,
  ReactionInfo
} from "./types.js";
import { log } from "./state.js";
import {
  getCapturedToken,
  getCapturedRegion,
  getCapturedTokenType,
  getLastSeenConversationId
} from "./worker-store.js";

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

const SERVICE_WORKER_RETRY_DELAY_MS = 500;
const SERVICE_WORKER_MAX_RETRIES = 3;

/**
 * Send a message to the background service worker with automatic retry.
 * MV3 service workers may not be immediately available after Chrome startup,
 * so we retry a few times on "Receiving end does not exist" errors.
 */
async function sendBackgroundMessageWithRetry<T>(message: Record<string, unknown>): Promise<T> {
  let lastError: Error | null = null;
  for (let attempt = 0; attempt <= SERVICE_WORKER_MAX_RETRIES; attempt++) {
    try {
      return await sendBackgroundMessage<T>(message);
    } catch (error) {
      lastError = error as Error;
      const isConnectionError = lastError.message.includes("Receiving end does not exist") ||
        lastError.message.includes("Could not establish connection");
      if (!isConnectionError || attempt === SERVICE_WORKER_MAX_RETRIES) {
        throw lastError;
      }
      log(`[api-client] Service worker not responding, retrying (${attempt + 1}/${SERVICE_WORKER_MAX_RETRIES})...`);
      await new Promise((resolve) => setTimeout(resolve, SERVICE_WORKER_RETRY_DELAY_MS * (attempt + 1)));
    }
  }
  throw lastError!;
}

// ── Direct fetch (content-script fallback) ────────────────────────────────────
//
// When the background service worker is unavailable (MV3 cold-start race, SW
// crash, or Playwright/debugging session quirk), the content script can make
// the API call directly.  Chrome grants cross-origin fetch to URLs listed in
// host_permissions, and the Teams page service worker only intercepts its own
// origin — not *.ng.msg.teams.microsoft.com.

const API_PAGE_SIZE = 200;
const API_MAX_PAGES = 100;

async function fetchChatMessagesDirectly(
  token: string,
  tokenType: string,
  region: string,
  conversationId: string
): Promise<ApiFetchResult> {
  const baseUrl = `https://${encodeURIComponent(region)}.ng.msg.teams.microsoft.com/v1`;
  const encodedConversationId = encodeURIComponent(conversationId);
  let nextUrl: string | null =
    `${baseUrl}/users/ME/conversations/${encodedConversationId}/messages?pageSize=${API_PAGE_SIZE}`;
  const allMessages: ApiMessage[] = [];
  let pageCount = 0;

  const authHeaders: Record<string, string> =
    tokenType === "bearer"
      ? { Authorization: `Bearer ${token}` }
      : { Authentication: `skypetoken=${token}` };

  while (nextUrl && pageCount < API_MAX_PAGES) {
    const response: Response = await fetch(nextUrl, { headers: authHeaders });

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

    const data: Record<string, unknown> = await response.json();

    if (Array.isArray(data.messages)) {
      allMessages.push(...(data.messages as ApiMessage[]));
    }

    nextUrl =
      data._metadata && (data._metadata as Record<string, unknown>).backwardLink
        ? ((data._metadata as Record<string, unknown>).backwardLink as string)
        : null;
    pageCount++;
  }

  return { messages: allMessages, pageCount };
}

// ── Public API ────────────────────────────────────────────────────────────────

export function isApiAvailable(): boolean {
  return Boolean(getCapturedToken() && getCapturedRegion());
}

export async function fetchFullChatViaApi(
  conversationId: string
): Promise<ApiFetchResult> {
  const token = getCapturedToken();
  const region = getCapturedRegion();

  if (!token || !region) {
    return { error: "No API credentials captured yet. Navigate to a Teams chat first." };
  }

  // Try background service worker first (preferred: avoids page SW interception)
  try {
    return await sendBackgroundMessageWithRetry<ApiFetchResult>({
      type: "teams-export/fetch-full-chat",
      conversationId,
      skypeToken: token,
      tokenType: getCapturedTokenType(),
      region
    });
  } catch (serviceWorkerError) {
    log(
      `[api-client] Background service worker unavailable (${(serviceWorkerError as Error).message}), falling back to direct fetch`
    );
  }

  // Fallback: fetch directly from the content script
  return fetchChatMessagesDirectly(
    token,
    getCapturedTokenType(),
    region,
    conversationId
  );
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
 * Extract the conversation ID from the DOM dataset attribute set by worker-store.
 * This is populated when the worker hook captures message batches with conversation IDs.
 */
function extractConversationIdFromDom(): string | null {
  const fromDataset = document.documentElement.dataset.teamsExportConversationId;
  if (fromDataset) return fromDataset;

  // Teams Cloud DOM elements contain the conversation ID in their IDs.
  // Chat headers:   chat-header-19:{uuid}_{uuid}@unq.gbl.spaces
  // Channel items:  sc-channel-list-item-19:{hex}@thread.tacv2
  const conversationIdPattern = /19:[a-f0-9-]+(?:_[a-f0-9-]+)?@(?:unq\.gbl\.spaces|thread\.(?:v2|tacv2|skype))/i;

  // Approach 1: find chat-title element and traverse upward to find chat-header
  const chatTitle = document.querySelector('[data-tid="chat-title"]');
  if (chatTitle) {
    let current: Element | null = chatTitle;
    for (let depth = 0; depth < 15 && current; depth++) {
      const elementId = current.id || "";
      if (elementId) {
        const match = elementId.match(conversationIdPattern);
        if (match) return match[0];
      }
      current = current.parentElement;
    }
  }

  // Approach 2: check chat-list-item elements for the selected/active one
  const chatListElements = document.querySelectorAll('[id*="chat-list-item_19:"]');
  for (const element of chatListElements) {
    const match = element.id.match(conversationIdPattern);
    if (!match) continue;

    let candidate: Element | null = element;
    for (let depth = 0; depth < 5 && candidate; depth++) {
      if (
        candidate.getAttribute("aria-selected") === "true" ||
        candidate.getAttribute("aria-current") === "true" ||
        candidate.classList.contains("active") ||
        candidate.classList.contains("selected")
      ) {
        return match[0];
      }
      candidate = candidate.parentElement;
    }
  }

  // Approach 3: check message pane runway parents for conversation ID
  const messagePane = document.querySelector('[data-tid="message-pane-list-runway"]');
  if (messagePane) {
    let current: Element | null = messagePane;
    for (let depth = 0; depth < 5 && current; depth++) {
      const elementId = current.id || "";
      if (elementId) {
        const match = elementId.match(conversationIdPattern);
        if (match) return match[0];
      }
      current = current.parentElement;
    }
  }

  // Approach 4 (fallback): scan chat-header elements
  const chatHeaderElements = document.querySelectorAll('[id*="chat-header-19:"]');
  for (const element of chatHeaderElements) {
    const match = element.id.match(conversationIdPattern);
    if (match) return match[0];
  }

  // Approach 5: active channel treeitem
  // In the Teams sidebar tree, the active/focused treeitem has tabindex="0".
  // Channel treeitems have data-testid like: favorite-channel-list-item-19:{hex}@thread.tacv2
  //                                   or: sc-channel-list-item-19:{hex}@thread.tacv2
  const activeChannelTreeItem = document.querySelector(
    '[role="treeitem"][tabindex="0"][data-testid*="channel-list-item-19:"]'
  );
  if (activeChannelTreeItem) {
    const testId = activeChannelTreeItem.getAttribute("data-testid") ?? "";
    const match = testId.match(conversationIdPattern);
    if (match) return match[0];
  }

  // Approach 6: channel sidebar — also check ID attribute of channel items
  // Some items have IDs like: title-channel-list-item-19:{hex}@thread.tacv2
  // This is a fallback when the treeitem approach above doesn't match.
  const channelTitleItems = document.querySelectorAll('[id*="channel-list-item-19:"]');
  for (const element of channelTitleItems) {
    const match = element.id.match(conversationIdPattern);
    if (match) return match[0];
  }

  return null;
}

/**
 * Get the conversation ID from the URL, DOM dataset, or from the worker's intercepted messages.
 */
export function resolveConversationId(): string | null {
  const fromUrl = extractConversationIdFromUrl();
  if (fromUrl) {
    return fromUrl;
  }

  const fromDom = extractConversationIdFromDom();
  if (fromDom) {
    return fromDom;
  }

  return getLastSeenConversationId();
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

// ── Channel thread detection ──────────────────────────────────────────────────

const CHANNEL_CONVERSATION_PATTERN = /@thread\.(tacv2|skype)$/i;
const THREAD_ID_PATTERN = /;messageid=(\d+)/;

function extractThreadId(conversationLink: string | undefined): string | null {
  if (!conversationLink) return null;
  const match = conversationLink.match(THREAD_ID_PATTERN);
  return match ? match[1] : null;
}

function isChannelConversation(apiMessages: ApiMessage[]): boolean {
  // Check if any message has a conversationLink with ;messageid= (channel thread pattern)
  return apiMessages.some(
    (message) => message.conversationLink && THREAD_ID_PATTERN.test(message.conversationLink)
  );
}

/**
 * Reorder channel messages so they are grouped by thread.
 * Each thread appears in order of its root post, with the root post first
 * followed by replies in chronological order.
 */
function groupChannelMessagesByThread(apiMessages: ApiMessage[]): ApiMessage[] {
  const threads = new Map<string, { root: ApiMessage | null; replies: ApiMessage[] }>();
  const threadOrder: string[] = [];

  for (const message of apiMessages) {
    const threadId = extractThreadId(message.conversationLink);
    if (!threadId) continue;

    if (!threads.has(threadId)) {
      threads.set(threadId, { root: null, replies: [] });
      threadOrder.push(threadId);
    }

    const thread = threads.get(threadId)!;
    if (message.id === threadId) {
      thread.root = message;
    } else {
      thread.replies.push(message);
    }
  }

  // Sort threads by root post time (oldest first)
  threadOrder.sort((threadIdA, threadIdB) => {
    const rootA = threads.get(threadIdA)!.root;
    const rootB = threads.get(threadIdB)!.root;
    const timeA = rootA?.originalarrivaltime ?? rootA?.composetime ?? "";
    const timeB = rootB?.originalarrivaltime ?? rootB?.composetime ?? "";
    return timeA.localeCompare(timeB);
  });

  const result: ApiMessage[] = [];
  for (const threadId of threadOrder) {
    const thread = threads.get(threadId)!;
    // Root post first
    if (thread.root) {
      result.push(thread.root);
    }
    // Replies in chronological order (API returns newest-first, so reverse)
    thread.replies.sort((replyA, replyB) => {
      const timeA = replyA.originalarrivaltime ?? replyA.composetime ?? "";
      const timeB = replyB.originalarrivaltime ?? replyB.composetime ?? "";
      return timeA.localeCompare(timeB);
    });
    result.push(...thread.replies);
  }

  return result;
}

// ── Snapshot conversion ───────────────────────────────────────────────────────

function convertSingleApiMessage(
  apiMessage: ApiMessage,
  index: number,
  threadId: string | null,
  isReply: boolean
): MessageSnapshot {
  const timestamp = formatApiTimestamp(
    apiMessage.originalarrivaltime ?? apiMessage.composetime
  );
  const content = apiMessage.content ?? "";
  const messageType = apiMessage.messagetype ?? apiMessage.type;

  return {
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
    captureOrder: index,
    ...(threadId ? { threadId, isReply } : {})
  };
}

export function convertApiMessagesToSnapshots(
  apiMessages: ApiMessage[]
): MessageSnapshot[] {
  const filteredMessages = apiMessages.filter(
    (message) => !EXCLUDED_MESSAGE_TYPES.has(message.messagetype ?? message.type ?? "")
  );

  const isChannel = isChannelConversation(filteredMessages);

  if (isChannel) {
    // Channel: group messages by thread (root post + replies)
    const grouped = groupChannelMessagesByThread(filteredMessages);
    return grouped.map((apiMessage, index) => {
      const threadId = extractThreadId(apiMessage.conversationLink);
      const isReply = threadId !== null && apiMessage.id !== threadId;
      return convertSingleApiMessage(apiMessage, index, threadId, isReply);
    });
  }

  // DM/group chat: simple chronological order
  const chronological = [...filteredMessages].reverse();
  return chronological.map((apiMessage, index) =>
    convertSingleApiMessage(apiMessage, index, null, false)
  );
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
