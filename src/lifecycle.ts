import {
  BODY_ACTIVE_CLASS,
  MESSAGE_CLASS,
  SELECTED_CLASS,
  CHECKBOX_CLASS,
  CONVERSATION_POLL_MS,
  MESSAGE_SCAN_INTERVAL_MS,
  INSTANCE_KEY
} from "./constants.js";
import { state, callbacks } from "./state.js";
import { isToolbarNode, isWithinMessageRunway } from "./utilities.js";
import { detectTheme } from "./theme.js";

export function applyTheme(): void {
  const nextTheme = detectTheme();
  state.theme = nextTheme;

  if (state.dock) {
    state.dock.dataset.theme = nextTheme;
  }
}

function handleSelectStart(event: Event): void {
  if (!state.active) {
    return;
  }

  if (isToolbarNode(event.target) || !isWithinMessageRunway(event.target)) {
    return;
  }

  event.preventDefault();
}

export function clearStartupTimers(): void {
  state.startupTimers.forEach((timerId) => window.clearTimeout(timerId));
  state.startupTimers = [];
}

export function scheduleStartupStabilization(): void {
  clearStartupTimers();

  for (const delay of [0, 40, 140, 320, 700, 1_400]) {
    const timerId: number = window.setTimeout(() => {
      callbacks.syncDockPosition();
      callbacks.applyTheme();
      callbacks.refreshMessages();
    }, delay);

    state.startupTimers.push(timerId);
  }
}

export function ensureGlobalHandlers(): void {
  if (!state.selectStartHandler) {
    state.selectStartHandler = handleSelectStart;
    document.addEventListener("selectstart", state.selectStartHandler, true);
  }

  if (!state.resizeHandler) {
    state.resizeHandler = () => {
      callbacks.syncDockPosition();
      callbacks.applyTheme();
    };
    window.addEventListener("resize", state.resizeHandler, { passive: true });
  }

  if (!state.themeMediaQuery && window.matchMedia) {
    state.themeMediaQuery = window.matchMedia("(prefers-color-scheme: dark)");
    state.themeMediaQuery.addEventListener?.("change", () => callbacks.applyTheme());
  }

  if (!state.focusHandler) {
    state.focusHandler = () => callbacks.checkConversationState();
    window.addEventListener("focus", state.focusHandler, { passive: true });
  }

  if (!state.visibilityHandler) {
    state.visibilityHandler = () => {
      if (document.visibilityState !== "hidden") {
        callbacks.checkConversationState();
      }
    };
    document.addEventListener("visibilitychange", state.visibilityHandler, {
      passive: true
    });
  }

  if (!state.pageShowHandler) {
    state.pageShowHandler = () => callbacks.checkConversationState();
    window.addEventListener("pageshow", state.pageShowHandler, { passive: true });
  }

  if (!state.popStateHandler) {
    state.popStateHandler = () => callbacks.checkConversationState();
    window.addEventListener("popstate", state.popStateHandler, { passive: true });
  }

  if (!state.hashChangeHandler) {
    state.hashChangeHandler = () => callbacks.checkConversationState();
    window.addEventListener("hashchange", state.hashChangeHandler, {
      passive: true
    });
  }

  if (!state.conversationPollTimer) {
    state.conversationPollTimer = window.setInterval((() => {
      if (document.visibilityState !== "hidden") {
        callbacks.checkConversationState();
      }
    }) as () => void, CONVERSATION_POLL_MS);
  }
}

/**
 * Start a periodic interval that scans for undecorated message rows and
 * triggers a refresh when new messages appear (e.g. after scrolling up to
 * load history).  Runs only while the panel is open or selection mode is active.
 */
export function startMessageScan(): void {
  if (state.messageScanTimer) {
    return;
  }

  state.messageScanTimer = window.setInterval(() => {
    if (document.visibilityState === "hidden") {
      return;
    }

    const undecoratedRows = document.querySelectorAll(
      `[data-tid="chat-pane-item"]:not(.${MESSAGE_CLASS}), ` +
      `[data-tid="channel-pane-message"]:not(.${MESSAGE_CLASS}), ` +
      `[data-tid="channel-replies-pane-message"]:not(.${MESSAGE_CLASS})`
    );

    const hasNewMessages = Array.from(undecoratedRows).some(
      (row) => (row.textContent || "").trim().length >= 4
    );

    if (hasNewMessages) {
      callbacks.refreshMessages();
    }
  }, MESSAGE_SCAN_INTERVAL_MS);
}

export function stopMessageScan(): void {
  if (state.messageScanTimer) {
    window.clearInterval(state.messageScanTimer);
    state.messageScanTimer = null;
  }
}

export function setupExtensionBridge(): void {
  const windowRecord = window as unknown as Record<string, unknown>;
  const globalRecord = globalThis as unknown as Record<string, unknown>;

  if (
    !globalRecord.chrome ||
    !(globalRecord.chrome as Record<string, unknown>).runtime ||
    windowRecord.__teamsMessageExporterBridgeBound
  ) {
    return;
  }

  const chromeRuntime = (globalRecord.chrome as { runtime?: { onMessage?: { addListener(callback: (message: unknown) => void): void } } }).runtime;

  if (!chromeRuntime?.onMessage) {
    return;
  }

  chromeRuntime.onMessage.addListener((message: unknown) => {
    if (
      !message ||
      typeof message !== "object" ||
      (message as Record<string, unknown>).type !== "teams-export/toggle-panel"
    ) {
      return;
    }

    callbacks.setPanelOpen(!state.panelOpen);
  });

  windowRecord.__teamsMessageExporterBridgeBound = true;
}

export function observeDom(): void {
  const observer = new MutationObserver((mutations) => {
    const shouldRefresh = mutations.some(
      (mutation) =>
        mutation.type === "childList" || mutation.type === "characterData"
    );
    const shouldApplyTheme = mutations.some(
      (mutation) =>
        mutation.type === "attributes" &&
        (mutation.target === document.documentElement ||
          mutation.target === document.body)
    );
    const shouldResyncDock = mutations.some(
      (mutation) =>
        mutation.type === "childList" ||
        mutation.target === document.documentElement ||
        mutation.target === document.body
    );

    if (shouldRefresh) {
      callbacks.debounceRefresh();
    }

    if (shouldApplyTheme) {
      callbacks.applyTheme();
    }

    if (shouldResyncDock) {
      callbacks.syncDockPosition();
    }
  });

  observer.observe(document.body, {
    subtree: true,
    childList: true,
    characterData: true
  });

  observer.observe(document.documentElement, {
    attributes: true,
    attributeFilter: ["class", "style", "data-theme", "data-app-theme"]
  });

  observer.observe(document.body, {
    attributes: true,
    attributeFilter: ["class", "style", "data-theme", "data-app-theme"]
  });

  state.observer = observer;
}

export function destroy(
  handleMessageMouseDown: (event: MouseEvent) => void,
  handleMessageClick: (event: MouseEvent) => void
): void {
  state.observer?.disconnect();
  window.clearTimeout(state.refreshTimer!);
  clearStartupTimers();

  if (state.selectStartHandler) {
    document.removeEventListener("selectstart", state.selectStartHandler, true);
  }

  if (state.resizeHandler) {
    window.removeEventListener("resize", state.resizeHandler);
  }

  if (state.focusHandler) {
    window.removeEventListener("focus", state.focusHandler);
    state.focusHandler = null;
  }

  if (state.visibilityHandler) {
    document.removeEventListener("visibilitychange", state.visibilityHandler);
    state.visibilityHandler = null;
  }

  if (state.pageShowHandler) {
    window.removeEventListener("pageshow", state.pageShowHandler);
    state.pageShowHandler = null;
  }

  if (state.popStateHandler) {
    window.removeEventListener("popstate", state.popStateHandler);
    state.popStateHandler = null;
  }

  if (state.hashChangeHandler) {
    window.removeEventListener("hashchange", state.hashChangeHandler);
    state.hashChangeHandler = null;
  }

  if (state.conversationPollTimer) {
    window.clearInterval(state.conversationPollTimer);
    state.conversationPollTimer = null;
  }

  stopMessageScan();

  if (state.themeMediaQuery) {
    state.themeMediaQuery = null;
  }

  document.body?.classList.remove(BODY_ACTIVE_CLASS);
  state.styleElement?.remove();
  state.dock?.remove();

  document.querySelectorAll(`.${MESSAGE_CLASS}`).forEach((element) => {
    element.classList.remove(MESSAGE_CLASS, SELECTED_CLASS);
    element
      .querySelectorAll(`:scope > .${CHECKBOX_CLASS}`)
      .forEach((node) => node.remove());
    element.removeEventListener(
      "mousedown",
      handleMessageMouseDown as EventListener,
      true
    );
    element.removeEventListener(
      "click",
      handleMessageClick as EventListener,
      true
    );
    delete (element as HTMLElement).dataset.tsmMessageId;
    delete (element as HTMLElement).dataset.tsmBound;
  });

  delete (window as unknown as Record<string, unknown>)[INSTANCE_KEY];
}
