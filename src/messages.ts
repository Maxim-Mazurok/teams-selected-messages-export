import type { Strategy, MessageRecord, MessageSnapshot, ReactionInfo } from "./types.js";
import { normalizeText, queryText } from "./utilities.js";
import {
  extractBodyHtml,
  extractPlainText,
  extractQuotedReply,
  extractReactions,
  extractPostReactions,
  extractSubject,
  getTimeMeta,
  isChannelPost
} from "./content-extraction.js";
import { elementToMarkdown } from "./markdown-renderer.js";
import { getMessageId, isLikelyMessageRow } from "./strategy.js";
import { log } from "./state.js";
import { getWorkerMessage, resolveUserId } from "./worker-store.js";

const THREAD_ROOT_ID_PATTERN = /^(?:reply-chain-summary-|content-|timestamp-)(\d+)$/;

/**
 * For a channel reply element, walk up the DOM to the parent
 * `channel-pane-message` root post and extract its thread root message ID.
 * Returns `{ threadId, isReply }` when the element is a reply, or
 * an empty object for root posts / non-channel messages.
 */
function extractThreadInfo(
  element: HTMLElement,
  messageId: string
): { threadId?: string; isReply?: boolean } {
  const responseSurface = element.closest('[data-tid="response-surface"]');
  if (!responseSurface) {
    return {};
  }

  const parentPost = responseSurface.closest('[data-tid="channel-pane-message"]');
  if (!parentPost) {
    return {};
  }

  const parentId = parentPost.id;
  const match = parentId.match(THREAD_ROOT_ID_PATTERN);
  if (!match) {
    return {};
  }

  const threadRootId = match[1];
  const normalizedMessageId = messageId.replace(
    /^(content-|timestamp-|reply-chain-summary-|message-body-)/,
    ""
  );

  if (normalizedMessageId === threadRootId) {
    return {};
  }

  return { threadId: threadRootId, isReply: true };
}

/**
 * Enrich DOM-extracted reactions with actor names resolved from worker data.
 *
 * The DOM reaction pill buttons expose reaction counts and emoji but typically
 * do not expose the full list of actors for each reaction.  The worker
 * interception bridge captures structured emotion objects that include every
 * reactor's user-ID; those IDs are resolved to display names using the member
 * map accumulated from author fields seen in worker responses.
 *
 * Matching strategy: the DOM reaction name (e.g. "like") is compared
 * case-insensitively against the worker emotion key ("like").  Teams keeps
 * these consistent so no emoji ↔ key lookup table is needed.
 */
function enrichReactionsFromWorker(messageId: string, reactions: ReactionInfo[]): ReactionInfo[] {
  const workerMessage = getWorkerMessage(messageId);
  if (!workerMessage) return reactions;

  const allEmotions = [...(workerMessage.emotions ?? []), ...(workerMessage.diverseEmotions ?? [])];
  if (allEmotions.length === 0) return reactions;

  return reactions.map((reaction) => {
    const normalizedName = reaction.name.toLowerCase().trim();
    const matchingEmotion = allEmotions.find(
      (emotion) => (emotion.key?.toLowerCase() ?? "") === normalizedName
    );

    if (!matchingEmotion || matchingEmotion.userIds.length === 0) return reaction;

    const resolvedActors = matchingEmotion.userIds
      .map((userId) => resolveUserId(userId))
      .filter((name): name is string => Boolean(name));

    // Fall back to DOM-extracted actors if we can't resolve the user IDs yet
    return resolvedActors.length > 0 ? { ...reaction, actors: resolvedActors } : reaction;
  });
}

export function buildMessageRecord(
  element: HTMLElement,
  index: number,
  strategy: Strategy | null
): MessageRecord | null {
  const plainText = extractPlainText(element, strategy);
  if (!plainText) {
    return null;
  }

  const timeMeta = getTimeMeta(element, strategy);
  const author =
    queryText(element, strategy?.authorSelector || "") ||
    queryText(element, '[aria-label*="sent by" i], [aria-label*="posted by" i]') ||
    "Unknown author";

  const messageId = getMessageId(element, index);
  const domReactions = isChannelPost(element) ? extractPostReactions(element) : extractReactions(element);

  const threadInfo = extractThreadInfo(element, messageId);

  return {
    id: messageId,
    index,
    element,
    author,
    timeLabel: timeMeta.label,
    dateTime: timeMeta.dateTime,
    subject: extractSubject(element),
    quote: extractQuotedReply(element),
    reactions: enrichReactionsFromWorker(messageId, domReactions),
    html: extractBodyHtml(element, strategy),
    markdown: elementToMarkdown(element, strategy, plainText),
    plainText,
    ...threadInfo
  };
}

export function snapshotMessageRecord(message: MessageRecord): MessageSnapshot {
  const { element: _element, ...snapshot } = message;
  return snapshot;
}

export function buildMessageRecordsFromRows(
  rows: Element[],
  strategy: Strategy | null
): MessageRecord[] {
  return rows.reduce<MessageRecord[]>((records, row, index) => {
    try {
      const record = buildMessageRecord(row as HTMLElement, index, strategy);
      if (record) {
        records.push(record);
      }
    } catch (error) {
      log("Skipping an unsupported Teams message row during refresh.", {
        index,
        messageId:
          row.getAttribute("data-message-id") ||
          row.id ||
          row.querySelector('[id^="content-"]')?.id ||
          null,
        error: error instanceof Error ? error.message : String(error)
      });
    }

    return records;
  }, []);
}

export function getVisibleMessageRecords(strategy: Strategy | null): MessageRecord[] {
  if (!strategy) {
    return [];
  }

  const rows = Array.from(document.querySelectorAll(strategy.rowSelector)).filter(isLikelyMessageRow);
  return buildMessageRecordsFromRows(rows, strategy);
}
