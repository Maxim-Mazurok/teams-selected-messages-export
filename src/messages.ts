import type { Strategy, MessageRecord, MessageSnapshot } from "./types.js";
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

  return {
    id: getMessageId(element, index),
    index,
    element,
    author,
    timeLabel: timeMeta.label,
    dateTime: timeMeta.dateTime,
    subject: extractSubject(element),
    quote: extractQuotedReply(element),
    reactions: isChannelPost(element) ? extractPostReactions(element) : extractReactions(element),
    html: extractBodyHtml(element, strategy),
    markdown: elementToMarkdown(element, strategy, plainText),
    plainText
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
