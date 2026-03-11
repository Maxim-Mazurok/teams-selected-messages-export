import type { MessageRecord, ExportPayload, ToggleSelectionOptions, ExportFullHistoryResult, MessageSnapshot } from "./types.js";
import {
  INSTANCE_KEY,
  BODY_ACTIVE_CLASS,
  MESSAGE_CLASS,
  SELECTED_CLASS,
  CHECKBOX_CLASS,
  HIDDEN_CHECKBOX_CLASS
} from "./constants.js";
import { state, callbacks, log } from "./state.js";
import { clearDomSelection, isToolbarNode, isInteractiveNode } from "./utilities.js";
import { generateStylesheet } from "./styles.js";
import { getConversationTitle, getConversationKey } from "./conversation.js";
import { selectStrategy } from "./strategy.js";
import { buildMessageRecordsFromRows, snapshotMessageRecord } from "./messages.js";
import { renderMarkdown } from "./markdown-renderer.js";
import { renderHtmlDocument } from "./html-renderer.js";
import { getMessageRange } from "./selection.js";
import {
  createExportPayload,
  commitExportPayload,
  writeClipboardText
} from "./export-helpers.js";
import { harvestFullChatMessages, exportFullHistory } from "./history.js";
import {
  createToolbar,
  updateToolbar,
  syncDockPosition,
  attachDock,
  setBusy
} from "./toolbar.js";
import {
  applyTheme,
  ensureGlobalHandlers,
  setupExtensionBridge,
  observeDom,
  scheduleStartupStabilization,
  clearStartupTimers,
  destroy
} from "./lifecycle.js";

/* ------------------------------------------------------------------ */
/*  Orchestration functions                                           */
/* ------------------------------------------------------------------ */

function debounceRefresh(): void {
  window.clearTimeout(state.refreshTimer!);
  state.refreshTimer = window.setTimeout(refreshMessages, 120) as unknown as number;
}

function hasDisconnectedMessages(): boolean {
  return (
    state.messages.length > 0 &&
    state.messages.every((message) => !message.element?.isConnected)
  );
}

function checkConversationState(): void {
  if (state.busy || !state.toolbar) {
    return;
  }

  const nextConversationKey = getConversationKey();
  const conversationChanged =
    Boolean(state.conversationKey) &&
    nextConversationKey !== state.conversationKey;

  if (conversationChanged || hasDisconnectedMessages()) {
    refreshMessages();
  }
}

function resetSelectionState(): void {
  state.selectedIds.clear();
  state.anchorId = null;
  state.lastPointerSelection = null;
  clearDomSelection();
}

function syncMessageDecorations(message: MessageRecord): void {
  message.element.classList.add(MESSAGE_CLASS);

  let wrapper = message.element.querySelector(
    `:scope > .${CHECKBOX_CLASS}`
  ) as HTMLLabelElement | null;

  if (!wrapper) {
    wrapper = document.createElement("label");
    wrapper.className = CHECKBOX_CLASS;
    wrapper.dataset.tsmSelectionControl = "true";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.setAttribute("aria-label", "Select Teams message");

    input.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleSelection(message.id, {
        shiftKey: Boolean(event.shiftKey),
        explicitValue: input.checked
      });
    });

    wrapper.appendChild(input);
    message.element.prepend(wrapper);
  }

  wrapper.classList.toggle(HIDDEN_CHECKBOX_CLASS, !state.active);
  message.element.classList.toggle(
    SELECTED_CLASS,
    state.active && state.selectedIds.has(message.id)
  );

  const input = wrapper.querySelector("input") as HTMLInputElement;
  input.checked = state.selectedIds.has(message.id);
}

function handleMessageMouseDown(event: MouseEvent): void {
  if (!state.active) {
    return;
  }

  const messageElement = event.currentTarget as HTMLElement;
  if (isToolbarNode(event.target) || isInteractiveNode(event.target)) {
    return;
  }

  const message = state.messageMap.get(messageElement.dataset.tsmMessageId!);
  if (!message) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  clearDomSelection();
  state.lastPointerSelection = { id: message.id, at: Date.now() };
  toggleSelection(message.id, { shiftKey: event.shiftKey });
}

function handleMessageClick(event: MouseEvent): void {
  if (!state.active) {
    return;
  }

  const messageElement = event.currentTarget as HTMLElement;
  if (isToolbarNode(event.target) || isInteractiveNode(event.target)) {
    return;
  }

  const message = state.messageMap.get(messageElement.dataset.tsmMessageId!);
  if (!message) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  if (
    state.lastPointerSelection &&
    state.lastPointerSelection.id === message.id &&
    Date.now() - state.lastPointerSelection.at < 400
  ) {
    return;
  }

  clearDomSelection();
  toggleSelection(message.id, { shiftKey: event.shiftKey });
}

function bindMessage(message: MessageRecord): void {
  message.element.dataset.tsmMessageId = message.id;

  if (!message.element.dataset.tsmBound) {
    message.element.dataset.tsmBound = "true";
    message.element.addEventListener(
      "mousedown",
      handleMessageMouseDown as EventListener,
      true
    );
    message.element.addEventListener(
      "click",
      handleMessageClick as EventListener,
      true
    );
  }

  syncMessageDecorations(message);
}

function cleanupOrphanedDecorations(): void {
  document.querySelectorAll(`.${MESSAGE_CLASS}`).forEach((element) => {
    if (!state.messages.some((message) => message.element === element)) {
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
    }
  });

  if (!state.busy) {
    for (const selectedId of Array.from(state.selectedIds)) {
      if (!state.messageMap.has(selectedId)) {
        state.selectedIds.delete(selectedId);
      }
    }
  }
}

function refreshMessages(): MessageRecord[] {
  const nextConversationKey = getConversationKey();
  const conversationChanged =
    Boolean(state.conversationKey) &&
    nextConversationKey !== state.conversationKey;
  state.conversationKey = nextConversationKey;

  if (conversationChanged) {
    resetSelectionState();
    scheduleStartupStabilization();
  }

  const { strategy, rows } = selectStrategy();
  state.strategy = strategy;
  state.messageMap.clear();

  const messages = buildMessageRecordsFromRows(rows, strategy);
  state.messages = messages;

  for (const message of messages) {
    state.messageMap.set(message.id, message);
    bindMessage(message);
  }

  cleanupOrphanedDecorations();
  syncDockPosition();
  updateToolbar();
  return messages;
}

function toggleSelection(
  messageId: string,
  options: ToggleSelectionOptions = {}
): void {
  const currentlySelected = state.selectedIds.has(messageId);
  const targetValue = options.explicitValue ?? !currentlySelected;

  if (options.shiftKey && state.anchorId) {
    const range = getMessageRange(
      state.messages.map((message) => message.id),
      state.anchorId,
      messageId
    );

    for (const id of range) {
      if (targetValue) {
        state.selectedIds.add(id);
      } else {
        state.selectedIds.delete(id);
      }
    }
  } else if (targetValue) {
    state.selectedIds.add(messageId);
  } else {
    state.selectedIds.delete(messageId);
  }

  state.anchorId = messageId;
  updateSelectionUi();
}

function clearSelection(): void {
  resetSelectionState();
  updateSelectionUi();
}

function setActive(active: boolean): void {
  state.active = active;
  document.body.classList.toggle(BODY_ACTIVE_CLASS, active);
  clearDomSelection();

  if (active && !state.messages.length) {
    refreshMessages();
  }

  updateToolbar();
  updateSelectionUi();
}

function setPanelOpen(open: boolean): void {
  state.panelOpen = open;

  if (open) {
    syncDockPosition();
    if (!state.messages.length) {
      refreshMessages();
    }
  }

  updateToolbar();
}

function updateSelectionUi(): void {
  state.messages.forEach(syncMessageDecorations);
  updateToolbar();
}

function getSelectedMessages(): MessageRecord[] {
  return state.messages.filter((message) => state.selectedIds.has(message.id));
}

function exportSelection(format: string): ExportPayload | null {
  const messages = getSelectedMessages().map(snapshotMessageRecord);
  if (!messages.length) {
    log("No selected messages to export.");
    return null;
  }

  const payload = commitExportPayload(
    createExportPayload(format, messages, { scope: "selection" })
  );
  state.lastExport = payload;
  return payload;
}

async function copyMarkdown(): Promise<boolean> {
  const messages = getSelectedMessages().map(snapshotMessageRecord);
  if (!messages.length) {
    return false;
  }

  const markdown = renderMarkdown(messages, {
    title: getConversationTitle(),
    sourceUrl: location.href,
    exportedAt: new Date().toISOString(),
    scope: "selection"
  });
  await writeClipboardText(markdown);
  state.lastExport = {
    format: "clipboard-md",
    content: markdown,
    count: messages.length,
    filename: "",
    scope: "selection"
  };
  setActive(false);
  setPanelOpen(false);
  return true;
}

/* ------------------------------------------------------------------ */
/*  Wire callbacks                                                    */
/* ------------------------------------------------------------------ */

callbacks.updateToolbar = updateToolbar;
callbacks.refreshMessages = refreshMessages;
callbacks.syncDockPosition = syncDockPosition;
callbacks.applyTheme = applyTheme;
callbacks.setBusy = setBusy;
callbacks.checkConversationState = checkConversationState;
callbacks.scheduleStartupStabilization = scheduleStartupStabilization;
callbacks.debounceRefresh = debounceRefresh;
callbacks.setActive = setActive;
callbacks.setPanelOpen = setPanelOpen;
callbacks.clearSelection = clearSelection;
callbacks.exportSelection = exportSelection;
callbacks.exportFullHistory = exportFullHistory;
callbacks.copyMarkdown = copyMarkdown;
callbacks.toggleSelection = toggleSelection;

/* ------------------------------------------------------------------ */
/*  Startup                                                           */
/* ------------------------------------------------------------------ */

function ensureStyles(): void {
  const styleElement = document.createElement("style");
  styleElement.dataset.tsmExport = "true";
  styleElement.textContent = generateStylesheet();
  document.head.appendChild(styleElement);
  state.styleElement = styleElement;
}

function registerApi(): void {
  const windowRecord = window as unknown as Record<string, unknown>;

  windowRecord[INSTANCE_KEY] = {
    state,
    refresh: refreshMessages,
    setActive,
    setPanelOpen,
    togglePanel: () => setPanelOpen(!state.panelOpen),
    clearSelection,
    exportSelection,
    exportFullHistory,
    renderHtml: () =>
      renderHtmlDocument(getSelectedMessages().map(snapshotMessageRecord), {
        title: getConversationTitle(),
        sourceUrl: location.href,
        exportedAt: new Date().toISOString()
      }),
    renderMarkdown: () =>
      renderMarkdown(getSelectedMessages().map(snapshotMessageRecord), {
        title: getConversationTitle(),
        sourceUrl: location.href,
        exportedAt: new Date().toISOString()
      }),
    getSelectedMessages: () =>
      getSelectedMessages().map(snapshotMessageRecord),
    destroy: () => destroy(handleMessageMouseDown, handleMessageClick)
  };
}

function start(): void {
  if (!document.head || !document.body) {
    return;
  }

  if (!state.toolbar) {
    ensureStyles();
    const elements = createToolbar({
      onClear: clearSelection,
      onExportHtml: () => exportSelection("html"),
      onExportMarkdown: () => exportSelection("md"),
      onExportFullMarkdown: () => {
        exportFullHistory("md").catch((error) => {
          log("Unable to export the full chat history.", error);
          setBusy(false);
        });
      },
      onExportFullHtml: () => {
        exportFullHistory("html").catch((error) => {
          log("Unable to export the full chat history.", error);
          setBusy(false);
        });
      },
      onCopyMarkdown: () => {
        copyMarkdown();
      },
      onTogglePanel: () => {
        if (state.panelOpen) {
          setActive(false);
          setPanelOpen(false);
          return;
        }

        setPanelOpen(true);
        if (!state.active) {
          setActive(true);
        }
      },
      onToggleActive: setActive
    });

    state.dock = elements.dock;
    state.toolbar = elements.toolbar;
    state.launcher = elements.launcher;
    state.quickCopyButton = elements.quickCopyButton;
    attachDock();
    applyTheme();
    syncDockPosition();
    scheduleStartupStabilization();
    ensureGlobalHandlers();
    setupExtensionBridge();
    observeDom();
  }

  applyTheme();
  refreshMessages();
  registerApi();
  log("Prototype loaded.");
}

/* ------------------------------------------------------------------ */
/*  Tear down previous instance and boot                              */
/* ------------------------------------------------------------------ */

const windowRecord = window as unknown as Record<string, unknown>;

if (
  windowRecord[INSTANCE_KEY] &&
  typeof (windowRecord[INSTANCE_KEY] as Record<string, unknown>).destroy === "function"
) {
  (windowRecord[INSTANCE_KEY] as { destroy(): void }).destroy();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", start, { once: true });
} else {
  start();
}
