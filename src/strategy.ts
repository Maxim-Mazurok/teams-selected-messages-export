import type { Strategy } from "./types.js";
import { STRATEGIES, RUNWAY_SELECTOR } from "./constants.js";
import { normalizeText, isToolbarNode } from "./utilities.js";

export function isLikelyMessageRow(element: Element): boolean {
  if (!element || isToolbarNode(element)) {
    return false;
  }

  if (!element.isConnected || !(element as HTMLElement).offsetParent) {
    return false;
  }

  const text = normalizeText(element.textContent || "");
  if (text.length < 4) {
    return false;
  }

  const isChannelPost = element.getAttribute("data-tid") === "channel-pane-message";
  if (isChannelPost) {
    return Boolean(element.querySelector('[data-tid="message-body"]'));
  }

  const isChannelRepliesPaneMessage =
    element.getAttribute("data-tid") === "channel-replies-pane-message";
  if (isChannelRepliesPaneMessage) {
    return Boolean(element.querySelector('[data-tid="message-body"]'));
  }

  const isChannelReplyGroup =
    element.getAttribute("role") === "group" &&
    Boolean(element.closest('[data-tid="response-surface"]')) &&
    Boolean(element.querySelector('[data-tid="message-body"]'));
  if (isChannelReplyGroup) {
    return true;
  }

  const hasStandardMessageShape = Boolean(
    element.querySelector('[data-testid="message-wrapper"]') &&
      element.querySelector(
        '[data-tid="chat-pane-message"], [data-message-content], [data-tid="message-body"], [data-tid="message-content"]'
      )
  );

  if (!hasStandardMessageShape) {
    return false;
  }

  const chatRunway = element.closest(RUNWAY_SELECTOR);
  if (chatRunway) {
    return true;
  }

  // Thread reply views can render message rows in side panes outside the classic runway container.
  return true;
}

export function scoreStrategy(strategy: Strategy): { rows: Element[]; score: number } {
  const rows = Array.from(document.querySelectorAll(strategy.rowSelector)).filter(isLikelyMessageRow);
  if (!rows.length) {
    return { rows: [], score: 0 };
  }

  const distinctTextRows = rows.filter((row) => normalizeText(row.textContent || "").length > 8);
  const score = distinctTextRows.length * 10 + Math.min(rows.length, 200);
  return { rows, score };
}

export function selectStrategy(): { strategy: Strategy | null; rows: Element[]; score: number } {
  let best: { strategy: Strategy | null; rows: Element[]; score: number } = {
    strategy: null,
    rows: [],
    score: 0
  };

  for (const strategy of STRATEGIES) {
    const result = scoreStrategy(strategy);
    if (result.score > best.score) {
      best = { strategy, rows: result.rows, score: result.score };
    }
  }

  return best;
}

export function getMessageId(element: Element, index: number): string {
  const explicitId =
    element.getAttribute("data-message-id") ||
    (element as HTMLElement).dataset.messageId ||
    element.getAttribute("data-item-id") ||
    element.id ||
    element.querySelector('[id^="content-"]')?.id ||
    element.querySelector('[id^="timestamp-"]')?.id;

  if (explicitId) {
    return explicitId;
  }

  const signature = normalizeText(element.textContent || "").slice(0, 120);
  return `derived-${index}-${signature}`;
}
