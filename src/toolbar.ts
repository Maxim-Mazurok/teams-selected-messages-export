import type { ToolbarActions, ToolbarElements } from "./types.js";
import {
  DOCK_CLASS,
  PANEL_CLASS,
  PANEL_OPEN_CLASS,
  LAUNCHER_ROW_CLASS,
  LAUNCHER_CLASS,
  QUICK_COPY_CLASS,
  HEADER_CLASS,
  STATUS_ROW_CLASS,
  PROGRESS_CLASS,
  TOGGLE_CLASS,
  COUNT_CLASS,
  ACTION_GRID_CLASS,
  ACTIVE_CLASS,
  HEADER_SLOT_SELECTORS,
  FALLBACK_DOCK_TOP,
  FALLBACK_DOCK_RIGHT
} from "./constants.js";
import { state, callbacks, log } from "./state.js";

function getDockHostElement(): Element | null {
  for (const selector of HEADER_SLOT_SELECTORS) {
    const element = document.querySelector(selector);
    if (element?.isConnected) {
      return element;
    }
  }

  return null;
}

export function attachDock(): void {
  if (!state.dock) {
    return;
  }

  const host = getDockHostElement();
  const nextParent = host || document.body;

  if (state.dock.parentElement !== nextParent) {
    nextParent.appendChild(state.dock);
  }
}

export function syncDockPosition(): void {
  if (!state.dock) {
    return;
  }

  attachDock();
  const host = getDockHostElement();

  if (!host) {
    state.dock.dataset.mode = "fallback";
    state.dock.style.top = `${FALLBACK_DOCK_TOP}px`;
    state.dock.style.right = `${FALLBACK_DOCK_RIGHT}px`;
    state.dock.style.bottom = "auto";
    state.dock.style.left = "auto";
    return;
  }

  state.dock.dataset.mode = "embedded";
  state.dock.style.top = "";
  state.dock.style.right = "";
  state.dock.style.bottom = "auto";
  state.dock.style.left = "";
}

export function createToolbar(actions: ToolbarActions): ToolbarElements {
  const dock = document.createElement("div");
  dock.className = DOCK_CLASS;

  const toolbar = document.createElement("aside");
  toolbar.className = PANEL_CLASS;

  const header = document.createElement("div");
  header.className = HEADER_CLASS;

  const title = document.createElement("h1");
  title.textContent = "Teams Selected Messages Export";

  const description = document.createElement("p");
  description.dataset.role = "description";
  description.textContent =
    "Select messages from the active conversation, then copy or export them cleanly.";

  const countLabel = document.createElement("span");
  countLabel.className = COUNT_CLASS;
  countLabel.dataset.role = "count";
  countLabel.textContent = "0 selected";

  header.append(title, description);

  const statusRow = document.createElement("div");
  statusRow.className = STATUS_ROW_CLASS;

  const selectionToggle = document.createElement("label");
  selectionToggle.className = TOGGLE_CLASS;

  const selectionInput = document.createElement("input");
  selectionInput.type = "checkbox";
  selectionInput.dataset.role = "active-toggle";

  const selectionText = document.createElement("span");
  selectionText.textContent = "Selection mode";

  selectionToggle.append(selectionInput, selectionText);
  statusRow.append(selectionToggle, countLabel);

  selectionInput.addEventListener("change", () => {
    actions.onToggleActive(selectionInput.checked);
  });

  const progress = document.createElement("p");
  progress.className = PROGRESS_CLASS;
  progress.dataset.role = "progress";
  progress.hidden = true;

  const actionGrid = document.createElement("div");
  actionGrid.className = ACTION_GRID_CLASS;
  actionGrid.dataset.role = "selection-actions";

  const copyButton = document.createElement("button");
  copyButton.type = "button";
  copyButton.dataset.action = "copy-md";
  copyButton.dataset.variant = "primary";
  copyButton.textContent = "Copy MD";

  const downloadMarkdownButton = document.createElement("button");
  downloadMarkdownButton.type = "button";
  downloadMarkdownButton.dataset.action = "export-md";
  downloadMarkdownButton.dataset.variant = "secondary";
  downloadMarkdownButton.textContent = "Download MD";

  const downloadHtmlButton = document.createElement("button");
  downloadHtmlButton.type = "button";
  downloadHtmlButton.dataset.action = "export-html";
  downloadHtmlButton.dataset.variant = "secondary";
  downloadHtmlButton.textContent = "Download HTML";

  const exportFullChatButton = document.createElement("button");
  exportFullChatButton.type = "button";
  exportFullChatButton.dataset.action = "export-full-md";
  exportFullChatButton.dataset.variant = "secondary";
  exportFullChatButton.textContent = "Full chat MD";

  const exportFullChatHtmlButton = document.createElement("button");
  exportFullChatHtmlButton.type = "button";
  exportFullChatHtmlButton.dataset.action = "export-full-html";
  exportFullChatHtmlButton.dataset.variant = "secondary";
  exportFullChatHtmlButton.textContent = "Full chat HTML";

  const clearButton = document.createElement("button");
  clearButton.type = "button";
  clearButton.dataset.action = "clear";
  clearButton.dataset.variant = "danger";
  clearButton.textContent = "Clear Selection";

  actionGrid.append(
    copyButton,
    clearButton,
    downloadMarkdownButton,
    downloadHtmlButton,
    exportFullChatButton,
    exportFullChatHtmlButton
  );

  toolbar.append(header, statusRow, progress, actionGrid);

  const launcherRow = document.createElement("div");
  launcherRow.className = LAUNCHER_ROW_CLASS;

  const quickCopyButton = document.createElement("button");
  quickCopyButton.className = QUICK_COPY_CLASS;
  quickCopyButton.type = "button";
  quickCopyButton.dataset.action = "copy-md";
  quickCopyButton.hidden = true;
  quickCopyButton.textContent = "Copy MD";

  const launcher = document.createElement("button");
  launcher.className = LAUNCHER_CLASS;
  launcher.type = "button";
  launcher.dataset.action = "toggle-panel";
  launcher.textContent = "Export";

  launcherRow.append(quickCopyButton, launcher);
  dock.append(launcherRow, toolbar);

  dock.addEventListener("click", (event) => {
    const button = (event.target as Element).closest(
      "button[data-action]"
    ) as HTMLButtonElement | null;
    if (!button) {
      return;
    }

    const action = button.dataset.action;

    if (action === "clear") {
      actions.onClear();
      return;
    }

    if (action === "export-html") {
      actions.onExportHtml();
      return;
    }

    if (action === "export-md") {
      actions.onExportMarkdown();
      return;
    }

    if (action === "export-full-md") {
      actions.onExportFullMarkdown();
      return;
    }

    if (action === "export-full-html") {
      actions.onExportFullHtml();
      return;
    }

    if (action === "copy-md") {
      actions.onCopyMarkdown();
      return;
    }

    if (action === "toggle-panel") {
      actions.onTogglePanel();
    }
  });

  return { dock, toolbar, launcher, quickCopyButton };
}

export function updateToolbar(): void {
  if (!state.toolbar) {
    return;
  }

  syncDockPosition();
  const selectedCount = state.selectedIds.size;
  const busy = state.busy;

  state.toolbar.classList.toggle(ACTIVE_CLASS, state.active);
  state.toolbar.classList.toggle(PANEL_OPEN_CLASS, state.panelOpen);

  const countElement = state.toolbar.querySelector('[data-role="count"]');
  if (countElement) {
    countElement.textContent = `${selectedCount} selected`;
  }

  const activeToggle = state.toolbar.querySelector(
    '[data-role="active-toggle"]'
  ) as HTMLInputElement | null;
  if (activeToggle) {
    activeToggle.checked = state.active;
    activeToggle.disabled = busy;
  }

  const progressElement = state.toolbar.querySelector(
    '[data-role="progress"]'
  ) as HTMLElement | null;
  if (progressElement) {
    progressElement.hidden = !busy;
    progressElement.textContent = state.busyText;
  }

  const actionButtons: Record<string, boolean> = {
    "export-html": busy || selectedCount === 0,
    "export-md": busy || selectedCount === 0,
    "export-full-md": busy,
    "export-full-html": busy,
    "copy-md": busy || !state.active || selectedCount === 0,
    clear: busy || selectedCount === 0
  };

  for (const [action, disabled] of Object.entries(actionButtons)) {
    const button = state.toolbar.querySelector(
      `[data-action="${action}"]`
    ) as HTMLButtonElement | null;
    if (button) {
      button.disabled = disabled;
    }
  }

  if (state.launcher) {
    state.launcher.textContent = "Export";
    state.launcher.setAttribute(
      "aria-expanded",
      state.panelOpen ? "true" : "false"
    );
    state.launcher.disabled = busy;
  }

  if (state.quickCopyButton) {
    state.quickCopyButton.hidden = !(selectedCount > 0 && !state.panelOpen);
    state.quickCopyButton.textContent =
      selectedCount > 0 ? `Copy ${selectedCount}` : "Copy MD";
    state.quickCopyButton.disabled = busy || selectedCount === 0;
  }
}

export function setBusy(active: boolean, text: string = ""): void {
  state.busy = active;
  state.busyText = text;
  callbacks.updateToolbar();
}
