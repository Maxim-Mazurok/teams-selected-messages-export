import type { Strategy, MessageSnapshot } from "./types.js";
import {
  RUNWAY_SELECTOR,
  HISTORY_LOADING_SELECTOR,
  HISTORY_POLL_INTERVAL_MS,
  HISTORY_SETTLE_MS,
  HISTORY_SETTLE_TOP_MS,
  HISTORY_MAX_WAIT_MS,
  HISTORY_MAX_WAIT_TOP_MS,
  HISTORY_TOP_STAGNANT_PASSES
} from "./constants.js";
import { getMessageId, isLikelyMessageRow } from "./strategy.js";
import { getVisibleMessageRecords, snapshotMessageRecord } from "./messages.js";
import { selectStrategy } from "./strategy.js";
import { orderMessagesForExport } from "./selection.js";
import { state, callbacks, log } from "./state.js";

function waitForDelay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function getRunwayElement(): Element | null {
  return document.querySelector(RUNWAY_SELECTOR);
}

export function isElementVisible(element: Element | null): boolean {
  if (!element || !element.isConnected) {
    return false;
  }

  const styles = window.getComputedStyle(element);
  if (styles.display === "none" || styles.visibility === "hidden" || styles.opacity === "0") {
    return false;
  }

  if (
    (element as HTMLElement).hidden ||
    element.getAttribute("aria-hidden") === "true"
  ) {
    return false;
  }

  const rect = element.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

export function isScrollableElement(element: Element | null): boolean {
  if (!element || element === document.body) {
    return false;
  }

  const styles = window.getComputedStyle(element);
  return (
    /(auto|scroll|overlay)/.test(styles.overflowY) &&
    element.scrollHeight > element.clientHeight + 8
  );
}

export function getMessageScrollContainer(): Element {
  const runway = getRunwayElement();
  let current: Element | null = runway;

  while (current) {
    if (isScrollableElement(current)) {
      return current;
    }

    current = current.parentElement;
  }

  return document.scrollingElement || document.documentElement;
}

function getVisibleHistoryRows(strategy: Strategy | null): Element[] {
  if (!strategy) {
    return [];
  }

  return Array.from(document.querySelectorAll(strategy.rowSelector)).filter(isLikelyMessageRow);
}

function getHistorySignature(
  messages: { id: string }[],
  scrollContainer: Element
): string {
  const firstId = messages[0]?.id || "none";
  const lastId = messages[messages.length - 1]?.id || "none";
  const top = Math.round(scrollContainer.scrollTop || 0);
  const height = Math.round(scrollContainer.scrollHeight || 0);
  return `${firstId}|${lastId}|${messages.length}|${top}|${height}`;
}

function getRowHistorySignature(rows: Element[], scrollContainer: Element): string {
  const firstId = rows[0] ? getMessageId(rows[0], 0) : "none";
  const lastId = rows.length
    ? getMessageId(rows[rows.length - 1], rows.length - 1)
    : "none";
  const top = Math.round(scrollContainer.scrollTop || 0);
  const height = Math.round(scrollContainer.scrollHeight || 0);
  return `${firstId}|${lastId}|${rows.length}|${top}|${height}`;
}

function hasActiveHistoryLoadingIndicator(
  strategy: Strategy | null,
  scrollContainer: Element
): boolean {
  const roots = [
    getRunwayElement(),
    scrollContainer,
    scrollContainer.parentElement,
    scrollContainer.closest("[role='main']")
  ];
  const rowSelector = strategy?.rowSelector || "";
  const seen = new Set<Element | null>();

  return roots.filter(Boolean).some((root) => {
    if (seen.has(root)) {
      return false;
    }

    seen.add(root);

    return Array.from(root!.querySelectorAll(HISTORY_LOADING_SELECTOR)).some(
      (candidate) => {
        if (!isElementVisible(candidate)) {
          return false;
        }

        if (rowSelector && candidate.closest(rowSelector)) {
          return false;
        }

        return true;
      }
    );
  });
}

async function waitForHistoryToSettle(
  strategy: Strategy | null,
  scrollContainer: Element
): Promise<string> {
  const nearTop = scrollContainer.scrollTop <= 4;
  const quietWindowMs = nearTop ? HISTORY_SETTLE_TOP_MS : HISTORY_SETTLE_MS;
  const maxWaitMs = nearTop ? HISTORY_MAX_WAIT_TOP_MS : HISTORY_MAX_WAIT_MS;
  const startedAt = performance.now();
  let lastChangeAt = startedAt;
  let lastSignature = getRowHistorySignature(
    getVisibleHistoryRows(strategy),
    scrollContainer
  );
  let lastLoadingState = hasActiveHistoryLoadingIndicator(strategy, scrollContainer);

  while (performance.now() - startedAt < maxWaitMs) {
    await waitForDelay(HISTORY_POLL_INTERVAL_MS);

    const currentSignature = getRowHistorySignature(
      getVisibleHistoryRows(strategy),
      scrollContainer
    );
    const loadingIndicatorVisible = hasActiveHistoryLoadingIndicator(
      strategy,
      scrollContainer
    );
    const now = performance.now();

    if (
      currentSignature !== lastSignature ||
      loadingIndicatorVisible !== lastLoadingState
    ) {
      lastChangeAt = now;
      lastSignature = currentSignature;
      lastLoadingState = loadingIndicatorVisible;
    }

    if (!loadingIndicatorVisible && now - lastChangeAt >= quietWindowMs) {
      break;
    }
  }

  return getRowHistorySignature(
    getVisibleHistoryRows(strategy),
    scrollContainer
  );
}

function collectVisibleMessagesIntoMap(
  strategy: Strategy | null,
  harvestedMap: Map<string, MessageSnapshot>,
  captureOrderStart: number = 0
): { visibleMessages: { id: string }[]; nextCaptureOrder: number } {
  const visibleMessages = getVisibleMessageRecords(strategy);
  let captureOrder = captureOrderStart;

  for (const message of visibleMessages) {
    if (!harvestedMap.has(message.id)) {
      harvestedMap.set(message.id, {
        ...snapshotMessageRecord(message),
        captureOrder
      });
      captureOrder += 1;
    }
  }

  return { visibleMessages, nextCaptureOrder: captureOrder };
}

export async function harvestFullChatMessages(): Promise<MessageSnapshot[]> {
  const strategy = state.strategy || selectStrategy().strategy;
  if (!strategy) {
    return [];
  }

  const scrollContainer = getMessageScrollContainer();
  const originalDistanceFromBottom =
    scrollContainer.scrollHeight - scrollContainer.scrollTop;
  const harvestedMap = new Map<string, MessageSnapshot>();
  let captureOrder = 0;
  let stagnantPasses = 0;
  let previousSignature = "";

  try {
    for (let pass = 0; pass < 160; pass += 1) {
      const beforeCollect = collectVisibleMessagesIntoMap(
        strategy,
        harvestedMap,
        captureOrder
      );
      captureOrder = beforeCollect.nextCaptureOrder;
      const sizeBeforeScroll = harvestedMap.size;
      callbacks.setBusy(
        true,
        `Loading full chat history... ${harvestedMap.size} messages captured`
      );

      const beforeTop = scrollContainer.scrollTop;
      const scrollStep = Math.max(
        260,
        Math.round(
          (scrollContainer.clientHeight || window.innerHeight || 800) * 0.82
        )
      );
      scrollContainer.scrollTop = Math.max(0, beforeTop - scrollStep);
      await waitForHistoryToSettle(strategy, scrollContainer);

      const afterCollect = collectVisibleMessagesIntoMap(
        strategy,
        harvestedMap,
        captureOrder
      );
      captureOrder = afterCollect.nextCaptureOrder;
      callbacks.setBusy(
        true,
        `Loading full chat history... ${harvestedMap.size} messages captured`
      );

      const currentSignature = getHistorySignature(
        afterCollect.visibleMessages,
        scrollContainer
      );
      const reachedTop = scrollContainer.scrollTop <= 4;
      const discoveredNewData =
        harvestedMap.size > sizeBeforeScroll ||
        currentSignature !== previousSignature;

      if (reachedTop && !discoveredNewData) {
        stagnantPasses += 1;
      } else {
        stagnantPasses = 0;
      }

      previousSignature = currentSignature;

      if (reachedTop && stagnantPasses >= HISTORY_TOP_STAGNANT_PASSES) {
        break;
      }
    }

    return orderMessagesForExport(Array.from(harvestedMap.values()));
  } finally {
    scrollContainer.scrollTop = Math.max(
      0,
      scrollContainer.scrollHeight - originalDistanceFromBottom
    );
    await waitForDelay(280);
    callbacks.refreshMessages();
  }
}

export async function exportFullHistory(
  format: string = "md",
  options: { download?: boolean; closePanel?: boolean } = {}
): Promise<
  | (import("./types.js").ExportPayload & { messages: MessageSnapshot[] })
  | null
> {
  if (state.busy) {
    return null;
  }

  const runway = getRunwayElement();
  if (!runway) {
    log("No Teams message runway found for full-history export.");
    return null;
  }

  callbacks.setBusy(true, "Loading full chat history...");

  try {
    const messages = await harvestFullChatMessages();
    if (!messages.length) {
      return null;
    }

    const { createExportPayload, commitExportPayload } = await import(
      "./export-helpers.js"
    );
    const payload = commitExportPayload(
      createExportPayload(format, messages, { scope: "full-chat" }),
      options
    );
    state.lastExport = payload;

    if (options.closePanel !== false) {
      callbacks.setPanelOpen(false);
    }

    return { ...payload, messages };
  } finally {
    callbacks.setBusy(false);
  }
}
