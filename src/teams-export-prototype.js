(() => {
  const INSTANCE_KEY = "__teamsMessageExporter";

  if (window[INSTANCE_KEY]?.destroy) {
    window[INSTANCE_KEY].destroy();
  }

  const UI_CLASS = "tsm-export";
  const DOCK_CLASS = `${UI_CLASS}__dock`;
  const PANEL_CLASS = `${UI_CLASS}__panel`;
  const PANEL_OPEN_CLASS = `${PANEL_CLASS}--open`;
  const LAUNCHER_ROW_CLASS = `${UI_CLASS}__launcher-row`;
  const LAUNCHER_CLASS = `${UI_CLASS}__launcher`;
  const QUICK_COPY_CLASS = `${UI_CLASS}__quick-copy`;
  const HEADER_CLASS = `${UI_CLASS}__header`;
  const STATUS_ROW_CLASS = `${UI_CLASS}__status-row`;
  const PROGRESS_CLASS = `${UI_CLASS}__progress`;
  const TOGGLE_CLASS = `${UI_CLASS}__toggle`;
  const COUNT_CLASS = `${UI_CLASS}__count`;
  const ACTION_GRID_CLASS = `${UI_CLASS}__action-grid`;
  const MESSAGE_CLASS = `${UI_CLASS}__message`;
  const SELECTED_CLASS = `${MESSAGE_CLASS}--selected`;
  const ACTIVE_CLASS = `${PANEL_CLASS}--active`;
  const CHECKBOX_CLASS = `${UI_CLASS}__checkbox`;
  const HIDDEN_CHECKBOX_CLASS = `${CHECKBOX_CLASS}--hidden`;
  const BODY_ACTIVE_CLASS = `${UI_CLASS}--selection-active`;
  const RUNWAY_SELECTOR = '[data-tid="message-pane-list-runway"], #chat-pane-list';
  const HISTORY_LOADING_SELECTOR = [
    '[role="progressbar"]',
    '[aria-busy="true"]',
    '[data-tid*="loading" i]',
    '[data-tid*="spinner" i]',
    '[data-tid*="loader" i]',
    '[data-testid*="loading" i]',
    '[data-testid*="spinner" i]',
    '[data-testid*="loader" i]'
  ].join(", ");
  const HISTORY_POLL_INTERVAL_MS = 120;
  const HISTORY_SETTLE_MS = 260;
  const HISTORY_SETTLE_TOP_MS = 720;
  const HISTORY_MAX_WAIT_MS = 1100;
  const HISTORY_MAX_WAIT_TOP_MS = 3200;
  const HISTORY_TOP_STAGNANT_PASSES = 2;
  const HEADER_SLOT_SELECTORS = [
    "#titlebar-end-slot-start",
    '[data-tid="titlebar-end-slot"]'
  ];
  const FALLBACK_DOCK_TOP = 10;
  const FALLBACK_DOCK_RIGHT = 200;
  const CONVERSATION_POLL_MS = 750;
  const CONTENT_PRUNE_SELECTOR = [
    `.${UI_CLASS}__toolbar`,
    `.${CHECKBOX_CLASS}`,
    '[role="menu"]',
    '[data-tid="quoted-reply-card"]',
    '[data-tid*="reactions"]',
    '[data-tid*="reaction"]',
    '[data-tid*="like"]',
    '[data-tid*="toolbar"]',
    '[data-tid*="avatar"] button',
    'button[aria-label*="reaction" i]',
    'button[aria-label*="options" i]',
    'button[aria-label*="more" i]'
  ].join(", ");

  const STRATEGIES = [
    {
      name: "teams-data-tid-chat-pane-item",
      rowSelector: '[data-tid="chat-pane-item"]',
      contentSelector: '[data-tid="chat-pane-message"], [data-message-content], [data-tid="message-body"], [data-tid="message-content"]',
      authorSelector: '[data-tid="message-author-name"], [data-tid="message-author"]',
      timeSelector: 'time, [data-tid="message-timestamp"], [data-tid="timestamp"]'
    },
    {
      name: "teams-data-tid-message-item",
      rowSelector: '[data-tid="message-pane-list-item"], [data-tid="message-item"], [data-tid="chat-message"]',
      contentSelector: '[data-tid="message-body"], [data-tid="message-content"], [data-tid="chat-pane-message"]',
      authorSelector: '[data-tid="message-author-name"], [data-tid="message-author"]',
      timeSelector: 'time, [data-tid="message-timestamp"], [data-tid="timestamp"]'
    },
    {
      name: "generic-listitem",
      rowSelector: 'div[role="listitem"], li[role="listitem"], [data-message-id]',
      contentSelector: '[data-tid="message-body"], [data-tid="message-content"], [dir="auto"], p, div',
      authorSelector: '[data-tid="message-author-name"], [data-tid="message-author"], [aria-label*="sent by" i]',
      timeSelector: 'time, [aria-label*=" at " i], [data-tid*="timestamp"]'
    }
  ];

  const state = {
    active: false,
    panelOpen: false,
    dock: null,
    toolbar: null,
    launcher: null,
    quickCopyButton: null,
    styleElement: null,
    observer: null,
    themeMediaQuery: null,
    refreshTimer: null,
    conversationPollTimer: null,
    startupTimers: [],
    selectStartHandler: null,
    resizeHandler: null,
    focusHandler: null,
    visibilityHandler: null,
    pageShowHandler: null,
    popStateHandler: null,
    hashChangeHandler: null,
    busy: false,
    busyText: "",
    messages: [],
    messageMap: new Map(),
    selectedIds: new Set(),
    anchorId: null,
    strategy: null,
    theme: "light",
    conversationKey: "",
    lastExport: null,
    lastPointerSelection: null
  };

  function log(...args) {
    console.log("[teams-export]", ...args);
  }

  function debounceRefresh() {
    window.clearTimeout(state.refreshTimer);
    state.refreshTimer = window.setTimeout(refreshMessages, 120);
  }

  function clearStartupTimers() {
    state.startupTimers.forEach((timerId) => window.clearTimeout(timerId));
    state.startupTimers = [];
  }

  function scheduleStartupStabilization() {
    clearStartupTimers();

    [0, 40, 140, 320, 700, 1400].forEach((delay) => {
      const timerId = window.setTimeout(() => {
        syncDockPosition();
        applyTheme();
        refreshMessages();
      }, delay);

      state.startupTimers.push(timerId);
    });
  }

  function ensureStyles() {
    const styleElement = document.createElement("style");
    styleElement.dataset.tsmExport = "true";
    styleElement.textContent = `
      .${DOCK_CLASS} {
        --tsm-panel-surface: rgba(248, 250, 255, 0.95);
        --tsm-panel-surface-strong: rgba(248, 250, 255, 0.92);
        --tsm-panel-surface-soft: rgba(13, 97, 216, 0.05);
        --tsm-panel-surface-muted: rgba(16, 35, 63, 0.06);
        --tsm-panel-surface-contrast: rgba(16, 35, 63, 0.08);
        --tsm-panel-border: rgba(13, 30, 54, 0.18);
        --tsm-panel-border-soft: rgba(13, 97, 216, 0.1);
        --tsm-panel-border-muted: rgba(16, 35, 63, 0.1);
        --tsm-panel-text: #10233f;
        --tsm-panel-text-strong: #173052;
        --tsm-panel-text-muted: #44556e;
        --tsm-panel-shadow: 0 14px 38px rgba(12, 23, 43, 0.18);
        --tsm-panel-highlight-shadow: 0 16px 42px rgba(0, 84, 190, 0.2);
        --tsm-button-primary-shadow: 0 10px 22px rgba(13, 97, 216, 0.2);
        --tsm-button-danger-surface: rgba(172, 32, 39, 0.08);
        --tsm-button-danger-border: rgba(172, 32, 39, 0.14);
        --tsm-button-danger-text: #9a232a;
        --tsm-selected-tint: rgba(13, 97, 216, 0.09);
        --tsm-selected-background:
          linear-gradient(90deg, rgba(0, 84, 190, 0.16), rgba(0, 84, 190, 0.06) 28%, rgba(0, 84, 190, 0.03));
        --tsm-selected-shadow:
          inset 0 0 0 1px rgba(0, 84, 190, 0.28),
          0 8px 18px rgba(0, 84, 190, 0.08);
        --tsm-checkbox-surface: rgba(248, 250, 255, 0.98);
        --tsm-checkbox-border: rgba(13, 30, 54, 0.1);
        --tsm-checkbox-shadow: 0 6px 14px rgba(12, 23, 43, 0.12);
        position: relative;
        display: flex;
        flex-direction: column;
        gap: 10px;
        align-items: flex-end;
        flex: 0 0 auto;
        min-width: 0;
        z-index: 2147483647;
      }

      .${DOCK_CLASS}[data-theme="dark"] {
        --tsm-panel-surface: rgba(19, 28, 43, 0.94);
        --tsm-panel-surface-strong: rgba(23, 34, 52, 0.92);
        --tsm-panel-surface-soft: rgba(86, 156, 255, 0.12);
        --tsm-panel-surface-muted: rgba(223, 232, 255, 0.08);
        --tsm-panel-surface-contrast: rgba(223, 232, 255, 0.12);
        --tsm-panel-border: rgba(173, 198, 255, 0.18);
        --tsm-panel-border-soft: rgba(86, 156, 255, 0.18);
        --tsm-panel-border-muted: rgba(223, 232, 255, 0.12);
        --tsm-panel-text: #ecf3ff;
        --tsm-panel-text-strong: #f8fbff;
        --tsm-panel-text-muted: #b8c6dd;
        --tsm-panel-shadow: 0 18px 42px rgba(4, 9, 18, 0.45);
        --tsm-panel-highlight-shadow: 0 18px 46px rgba(12, 30, 66, 0.48);
        --tsm-button-primary-shadow: 0 10px 24px rgba(24, 106, 224, 0.32);
        --tsm-button-danger-surface: rgba(255, 120, 120, 0.12);
        --tsm-button-danger-border: rgba(255, 145, 145, 0.18);
        --tsm-button-danger-text: #ffb0b0;
        --tsm-selected-tint: rgba(66, 145, 255, 0.12);
        --tsm-selected-background:
          linear-gradient(90deg, rgba(66, 145, 255, 0.22), rgba(66, 145, 255, 0.1) 30%, rgba(66, 145, 255, 0.04));
        --tsm-selected-shadow:
          inset 0 0 0 1px rgba(104, 166, 255, 0.38),
          0 10px 24px rgba(8, 18, 34, 0.28);
        --tsm-checkbox-surface: rgba(16, 24, 37, 0.98);
        --tsm-checkbox-border: rgba(223, 232, 255, 0.18);
        --tsm-checkbox-shadow: 0 8px 18px rgba(4, 9, 18, 0.34);
      }

      .${DOCK_CLASS}[data-mode="embedded"] {
        margin-right: 12px;
      }

      .${DOCK_CLASS}[data-mode="fallback"] {
        position: fixed;
        top: 10px;
        right: 200px;
      }

      .${PANEL_CLASS} {
        position: absolute;
        top: calc(100% + 10px);
        right: 0;
        width: 340px;
        max-width: min(340px, calc(100vw - 24px));
        padding: 14px;
        border-radius: 16px;
        border: 1px solid var(--tsm-panel-border);
        background: var(--tsm-panel-surface);
        box-shadow: var(--tsm-panel-shadow);
        backdrop-filter: blur(18px);
        color: var(--tsm-panel-text);
        font: 13px/1.35 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        white-space: normal;
        opacity: 0;
        transform: translateY(-8px) scale(0.98);
        transform-origin: top right;
        pointer-events: none;
        transition: opacity 140ms ease, transform 160ms ease;
      }

      .${PANEL_CLASS}.${PANEL_OPEN_CLASS} {
        opacity: 1;
        transform: translateY(0) scale(1);
        pointer-events: auto;
      }

      .${ACTIVE_CLASS} {
        border-color: color-mix(in srgb, var(--tsm-panel-border) 20%, #0d61d8);
        box-shadow: var(--tsm-panel-highlight-shadow);
      }

      .${LAUNCHER_ROW_CLASS} {
        display: flex;
        align-items: center;
        gap: 8px;
        justify-content: flex-end;
        min-height: 36px;
      }

      .${LAUNCHER_CLASS},
      .${QUICK_COPY_CLASS} {
        appearance: none;
        border: 1px solid var(--tsm-panel-border-muted);
        border-radius: 999px;
        background: var(--tsm-panel-surface-strong);
        color: var(--tsm-panel-text-strong);
        font: 12px/1 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        box-shadow: 0 8px 18px color-mix(in srgb, var(--tsm-panel-text) 12%, transparent);
        backdrop-filter: blur(16px);
        cursor: pointer;
      }

      .${LAUNCHER_CLASS} {
        display: inline-flex;
        align-items: center;
        gap: 9px;
        min-height: 34px;
        padding: 0 12px 0 9px;
        font-weight: 700;
      }

      .${LAUNCHER_CLASS}::before {
        content: "";
        width: 10px;
        height: 10px;
        border-radius: 999px;
        background: radial-gradient(circle at 30% 30%, #78afff, #0d61d8 72%);
        box-shadow: 0 0 0 3px rgba(13, 97, 216, 0.1);
      }

      .${QUICK_COPY_CLASS} {
        min-height: 32px;
        padding: 0 11px;
        font-weight: 700;
      }

      .${QUICK_COPY_CLASS}[hidden] {
        display: none;
      }

      .${LAUNCHER_CLASS}:disabled,
      .${QUICK_COPY_CLASS}:disabled {
        opacity: 0.5;
        cursor: not-allowed;
        box-shadow: none;
      }

      .${HEADER_CLASS} {
        display: grid;
        gap: 8px;
        min-width: 0;
        margin-bottom: 14px;
      }

      .${PANEL_CLASS} h1 {
        margin: 0;
        font-size: 14px;
        font-weight: 700;
      }

      .${PANEL_CLASS} p {
        margin: 0;
        color: var(--tsm-panel-text-muted);
        line-height: 1.45;
        white-space: normal;
        overflow-wrap: anywhere;
      }

      .${STATUS_ROW_CLASS} {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 14px;
        padding: 10px 12px;
        border-radius: 14px;
        background: var(--tsm-panel-surface-soft);
        border: 1px solid var(--tsm-panel-border-soft);
      }

      .${PROGRESS_CLASS} {
        margin: 0 0 12px;
        color: var(--tsm-panel-text-muted);
        font-size: 12px;
        line-height: 1.4;
      }

      .${PROGRESS_CLASS}[hidden] {
        display: none;
      }

      .${TOGGLE_CLASS} {
        display: inline-flex;
        align-items: center;
        gap: 10px;
        color: var(--tsm-panel-text-strong);
        font-size: 12px;
        font-weight: 700;
      }

      .${TOGGLE_CLASS} input {
        width: 16px;
        height: 16px;
        margin: 0;
        accent-color: #0d61d8;
      }

      .${COUNT_CLASS} {
        display: inline-flex;
        align-items: center;
        min-height: 28px;
        padding: 0 10px;
        border-radius: 999px;
        background: var(--tsm-panel-surface-contrast);
        color: var(--tsm-panel-text-strong);
        font-size: 12px;
        font-weight: 700;
        white-space: nowrap;
      }

      .${ACTION_GRID_CLASS} {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 8px;
        margin-bottom: 8px;
        align-items: stretch;
      }

      .${ACTION_GRID_CLASS} > button {
        width: 100%;
        gap: 8px;
      }

      .${PANEL_CLASS} button {
        appearance: none;
        border: 1px solid transparent;
        border-radius: 12px;
        min-height: 42px;
        padding: 10px 12px;
        font: inherit;
        font-weight: 700;
        cursor: pointer;
        transition: transform 120ms ease, box-shadow 120ms ease, background-color 120ms ease;
      }

      .${PANEL_CLASS} button:hover:not(:disabled) {
        transform: translateY(-1px);
      }

      .${PANEL_CLASS} button[data-variant="primary"] {
        background: linear-gradient(135deg, #0d61d8, #0f4ea6);
        color: #fff;
        box-shadow: var(--tsm-button-primary-shadow);
      }

      .${PANEL_CLASS} button[data-variant="secondary"] {
        background: var(--tsm-panel-surface-muted);
        color: var(--tsm-panel-text-strong);
        border-color: var(--tsm-panel-border-muted);
      }

      .${PANEL_CLASS} button[data-variant="danger"] {
        background: var(--tsm-button-danger-surface);
        color: var(--tsm-button-danger-text);
        border-color: var(--tsm-button-danger-border);
      }

      .${PANEL_CLASS} button:disabled {
        opacity: 0.45;
        cursor: not-allowed;
        transform: none;
        box-shadow: none;
      }

      .${MESSAGE_CLASS} {
        position: relative;
        border-radius: 18px;
        transition: background-color 120ms ease, box-shadow 120ms ease;
      }

      .${MESSAGE_CLASS}.${SELECTED_CLASS} {
        background-color: var(--tsm-selected-tint);
        background-image: var(--tsm-selected-background);
        box-shadow: var(--tsm-selected-shadow);
      }

      .${MESSAGE_CLASS}.${SELECTED_CLASS}::before {
        content: "";
        position: absolute;
        top: 8px;
        bottom: 8px;
        left: 6px;
        width: 4px;
        border-radius: 999px;
        background: linear-gradient(180deg, #0d61d8, #3c8dff);
      }

      .${CHECKBOX_CLASS} {
        position: absolute;
        top: 10px;
        left: -30px;
        z-index: 5;
        display: flex;
        align-items: center;
        justify-content: center;
        width: 22px;
        height: 22px;
        border-radius: 999px;
        background: var(--tsm-checkbox-surface);
        border: 1px solid var(--tsm-checkbox-border);
        box-shadow: var(--tsm-checkbox-shadow);
      }

      .${CHECKBOX_CLASS} input {
        width: 14px;
        height: 14px;
        margin: 0;
      }

      .${CHECKBOX_CLASS}.${HIDDEN_CHECKBOX_CLASS} {
        display: none;
      }

      body.${BODY_ACTIVE_CLASS} ${RUNWAY_SELECTOR},
      body.${BODY_ACTIVE_CLASS} ${RUNWAY_SELECTOR} * {
        user-select: none;
        -webkit-user-select: none;
      }
    `;

    document.head.appendChild(styleElement);
    state.styleElement = styleElement;
  }

  function createToolbar() {
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
    description.textContent = "Select messages from the active conversation, then copy or export them cleanly.";

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
      setActive(selectionInput.checked);
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
      const button = event.target.closest("button[data-action]");
      if (!button) {
        return;
      }

      const action = button.dataset.action;

      if (action === "clear") {
        clearSelection();
        return;
      }

      if (action === "export-html") {
        exportSelection("html");
        return;
      }

      if (action === "export-md") {
        exportSelection("md");
        return;
      }

      if (action === "export-full-md") {
        exportFullHistory("md").catch((error) => {
          log("Unable to export the full chat history.", error);
          setBusy(false);
        });
        return;
      }

      if (action === "export-full-html") {
        exportFullHistory("html").catch((error) => {
          log("Unable to export the full chat history.", error);
          setBusy(false);
        });
        return;
      }

      if (action === "copy-md") {
        copyMarkdown();
        return;
      }

      if (action === "toggle-panel") {
        if (state.panelOpen) {
          setPanelOpen(false);
          return;
        }

        setPanelOpen(true);
        if (!state.active) {
          setActive(true);
        }
      }
    });

    state.dock = dock;
    state.toolbar = toolbar;
    state.launcher = launcher;
    state.quickCopyButton = quickCopyButton;
    attachDock();
    applyTheme();
    syncDockPosition();
    scheduleStartupStabilization();
  }

  function isToolbarNode(node) {
    return Boolean(node?.closest?.(`.${DOCK_CLASS}`));
  }

  function isInteractiveNode(node) {
    return Boolean(
      node?.closest?.(
        'a, button, input, textarea, summary, select, option, [contenteditable="true"], [data-tsm-selection-control="true"]'
      )
    );
  }

  function isWithinMessageRunway(node) {
    return Boolean(node?.closest?.(RUNWAY_SELECTOR));
  }

  function isMentionElement(element) {
    const ariaLabel = element?.getAttribute?.("aria-label") || "";
    return (
      ariaLabel.startsWith("Mentioned ") ||
      element?.getAttribute?.("itemtype") === "http://schema.skype.com/Mention"
    );
  }

  function clearDomSelection() {
    const selection = window.getSelection?.();
    if (selection && selection.rangeCount) {
      selection.removeAllRanges();
    }
  }

  function parseRgbColor(value) {
    const match = String(value || "")
      .trim()
      .match(/^rgba?\(([^)]+)\)$/i);
    if (!match) {
      return null;
    }

    const channels = match[1]
      .split(",")
      .map((part) => Number.parseFloat(part.trim()))
      .filter((part) => Number.isFinite(part));

    const [red = 0, green = 0, blue = 0, alpha = 1] = channels;
    return { red, green, blue, alpha };
  }

  function relativeLuminanceChannel(channel) {
    const normalized = channel / 255;
    if (normalized <= 0.03928) {
      return normalized / 12.92;
    }

    return ((normalized + 0.055) / 1.055) ** 2.4;
  }

  function isDarkRgb(color) {
    if (!color || color.alpha === 0) {
      return null;
    }

    const luminance =
      0.2126 * relativeLuminanceChannel(color.red) +
      0.7152 * relativeLuminanceChannel(color.green) +
      0.0722 * relativeLuminanceChannel(color.blue);

    return luminance < 0.35;
  }

  function getTeamsThemeHint() {
    const classNames = [document.documentElement.className, document.body?.className || ""].join(" ");

    if (/\btheme-(dark|black|contrast)(?:[a-z0-9-]*)\b|\bdark\b|\bcontrast\b/i.test(classNames)) {
      return "dark";
    }

    if (/\btheme-(default|light)(?:[a-z0-9-]*)\b|\blight\b/i.test(classNames)) {
      return "light";
    }

    return null;
  }

  function getPageThemeFallback() {
    const probes = [
      document.querySelector('[data-tid="title-bar"]'),
      document.querySelector('[data-tid="titlebar-end-slot"]'),
      document.body,
      document.documentElement
    ].filter(Boolean);

    for (const probe of probes) {
      const backgroundColor = parseRgbColor(window.getComputedStyle(probe).backgroundColor);
      const isDark = isDarkRgb(backgroundColor);
      if (isDark !== null) {
        return isDark ? "dark" : "light";
      }
    }

    return window.matchMedia?.("(prefers-color-scheme: dark)")?.matches ? "dark" : "light";
  }

  function detectTheme() {
    return getTeamsThemeHint() || getPageThemeFallback();
  }

  function applyTheme() {
    const nextTheme = detectTheme();
    state.theme = nextTheme;

    if (state.dock) {
      state.dock.dataset.theme = nextTheme;
    }
  }

  function getDockHostElement() {
    for (const selector of HEADER_SLOT_SELECTORS) {
      const element = document.querySelector(selector);
      if (element?.isConnected) {
        return element;
      }
    }

    return null;
  }

  function attachDock() {
    if (!state.dock) {
      return;
    }

    const host = getDockHostElement();
    const nextParent = host || document.body;

    if (state.dock.parentElement !== nextParent) {
      nextParent.appendChild(state.dock);
    }
  }

  function syncDockPosition() {
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

  function getConversationTitle() {
    const candidates = [
      '[data-tid="chat-title"]',
      '[data-tid="chat-header-title"]',
      '[data-tid="thread-header-title"]',
      'main h1',
      '[role="main"] h1',
      '[aria-level="1"]'
    ];

    for (const selector of candidates) {
      const element = document.querySelector(selector);
      const text = normalizeText(element?.textContent || "");
      if (text) {
        return text;
      }
    }

    return document.title.replace(/\s*-\s*Microsoft Teams\s*$/i, "").trim() || "Microsoft Teams Conversation";
  }

  function getConversationKey() {
    const title = getConversationTitle();
    const documentTitle = normalizeText(document.title || "");
    const path = `${location.pathname}${location.search}${location.hash}`;
    return [title, documentTitle, path].filter(Boolean).join(" | ");
  }

  function normalizeText(value) {
    return value.replace(/\s+/g, " ").trim();
  }

  function escapeHtml(value) {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function formatLocalDateTime(value) {
    const date = value instanceof Date ? value : new Date(value);
    return new Intl.DateTimeFormat(undefined, {
      dateStyle: "long",
      timeStyle: "medium"
    }).format(date);
  }

  function filenameSafe(value) {
    return value.replace(/[^a-z0-9._-]+/gi, "-").replace(/-+/g, "-").replace(/^-|-$/g, "").toLowerCase();
  }

  function queryText(root, selectorList) {
    if (!selectorList) {
      return "";
    }

    const selectors = selectorList
      .split(",")
      .map((entry) => entry.trim())
      .filter(Boolean);

    for (const selector of selectors) {
      const text = normalizeText(root.querySelector(selector)?.textContent || "");
      if (text) {
        return text;
      }
    }

    return "";
  }

  function scoreStrategy(strategy) {
    const rows = Array.from(document.querySelectorAll(strategy.rowSelector)).filter(isLikelyMessageRow);
    if (!rows.length) {
      return { rows: [], score: 0 };
    }

    const distinctTextRows = rows.filter((row) => normalizeText(row.textContent || "").length > 8);
    const score = distinctTextRows.length * 10 + Math.min(rows.length, 200);
    return { rows, score };
  }

  function isLikelyMessageRow(element) {
    if (!element || isToolbarNode(element)) {
      return false;
    }

    if (!element.isConnected || !element.offsetParent) {
      return false;
    }

    const text = normalizeText(element.textContent || "");
    if (text.length < 4) {
      return false;
    }

    const chatRunway = element.closest(RUNWAY_SELECTOR);
    if (!chatRunway) {
      return false;
    }

    return Boolean(
      element.querySelector('[data-testid="message-wrapper"]') &&
        element.querySelector('[data-tid="chat-pane-message"], [data-message-content]')
    );
  }

  function selectStrategy() {
    let best = { strategy: null, rows: [], score: 0 };

    for (const strategy of STRATEGIES) {
      const result = scoreStrategy(strategy);
      if (result.score > best.score) {
        best = { strategy, rows: result.rows, score: result.score };
      }
    }

    return best;
  }

  function getMessageId(element, index) {
    const explicitId =
      element.getAttribute("data-message-id") ||
      element.dataset.messageId ||
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

  function pruneClone(root) {
    root.querySelectorAll(CONTENT_PRUNE_SELECTOR).forEach((node) => node.remove());
    return root;
  }

  function getContentSource(element, strategy) {
    return (strategy?.contentSelector && element.querySelector(strategy.contentSelector)) || element;
  }

  function formatMentionLabel(element) {
    const ariaLabel = element.getAttribute("aria-label") || "";
    const explicitName = ariaLabel.startsWith("Mentioned ") ? ariaLabel.replace(/^Mentioned\s+/, "") : "";
    const text = normalizeText(element.textContent || "");
    const name = explicitName || text;
    return name ? `@${name}` : text;
  }

  function isEmojiImage(image) {
    return Boolean(
      image.closest('[data-tid="emoticon-renderer"]') ||
        image.getAttribute("itemtype") === "http://schema.skype.com/Emoji"
    );
  }

  function isDecorativeImage(image) {
    return Boolean(
      image.closest('[data-tid*="reaction"]') ||
        image.closest('[data-tid="message-avatar"]') ||
        image.closest('[data-tid="quoted-reply-card"]')
    );
  }

  function createImagePlaceholder(label) {
    const placeholder = document.createElement("span");
    placeholder.textContent = label;
    placeholder.dataset.tsmPlaceholder = "image";
    return placeholder;
  }

  function normalizeContentClone(root) {
    root.querySelectorAll("*").forEach((element) => {
      if (isMentionElement(element)) {
        const replacement = document.createElement("span");
        replacement.textContent = formatMentionLabel(element);
        replacement.dataset.tsmMention = "true";
        element.replaceWith(replacement);
        return;
      }

      if (element.tagName?.toLowerCase() !== "img") {
        return;
      }

      if (isEmojiImage(element)) {
        const replacement = document.createTextNode(element.getAttribute("alt") || "");
        element.replaceWith(replacement);
        return;
      }

      if (isDecorativeImage(element)) {
        element.remove();
        return;
      }

      element.replaceWith(createImagePlaceholder("[Image omitted]"));
    });

    return root;
  }

  function getPreparedContentClone(element, strategy) {
    const clone = pruneClone(getContentSource(element, strategy).cloneNode(true));
    return normalizeContentClone(clone);
  }

  function extractBodyHtml(element, strategy) {
    const clone = getPreparedContentClone(element, strategy);
    return clone.innerHTML.trim();
  }

  function extractPlainText(element, strategy) {
    return normalizeText(getPreparedContentClone(element, strategy).textContent || "");
  }

  function extractQuotedReply(element) {
    const card = element.querySelector('[data-tid="quoted-reply-card"]');
    if (!card) {
      return null;
    }

    const timeLabel = normalizeText(card.querySelector('[data-tid="quoted-reply-timestamp"]')?.textContent || "");
    const previewText =
      normalizeText(card.querySelector('[data-tid="quoted-reply-preview-content"]')?.textContent || "") ||
      normalizeText(card.textContent || "");

    const author = Array.from(card.querySelectorAll("span"))
      .map((node) => normalizeText(node.textContent || ""))
      .find((text) => text && text !== timeLabel && text !== previewText);

    return {
      author: author || "",
      timeLabel,
      text: previewText
    };
  }

  function extractReactionActors(button, labelText) {
    const candidates = [
      button.getAttribute("aria-label"),
      button.getAttribute("title"),
      labelText
    ]
      .map((value) => normalizeText(value || ""))
      .filter(Boolean);

    const actors = [];

    candidates.forEach((candidate) => {
      const byMatch = candidate.match(/(?:from|by)\s+(.+)$/i);
      if (!byMatch) {
        return;
      }

      byMatch[1]
        .split(/,|&| and /i)
        .map((part) => normalizeText(part))
        .filter(Boolean)
        .forEach((name) => actors.push(name));
    });

    return Array.from(new Set(actors));
  }

  function extractReactions(element) {
    return Array.from(element.querySelectorAll('[data-tid="diverse-reaction-pill-button"]'))
      .map((button) => {
        const labelId = button.getAttribute("aria-labelledby");
        const labelText = normalizeText(
          (labelId && document.getElementById(labelId)?.textContent) || button.textContent || ""
        );
        const match = labelText.match(/(\d+)\s+(.+?)\s+reactions?\.?/i);
        const emoji = button.querySelector('img[alt]')?.getAttribute("alt") || "";
        const count = Number(match?.[1] || 1);
        const name = normalizeText(match?.[2] || labelText.replace(/\.$/, ""));
        const actors = extractReactionActors(button, labelText);

        if (!emoji && !name) {
          return null;
        }

        return {
          emoji,
          name,
          count,
          actors
        };
      })
      .filter(Boolean);
  }

  function getTimeMeta(element, strategy) {
    const timeElement = strategy?.timeSelector ? element.querySelector(strategy.timeSelector) : element.querySelector("time");

    if (!timeElement) {
      return { label: "", dateTime: "" };
    }

    return {
      label: normalizeText(timeElement.textContent || ""),
      dateTime: timeElement.getAttribute("datetime") || timeElement.dateTime || ""
    };
  }

  function buildMessageRecord(element, index, strategy) {
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
      quote: extractQuotedReply(element),
      reactions: extractReactions(element),
      html: extractBodyHtml(element, strategy),
      markdown: elementToMarkdown(element, strategy, plainText),
      plainText
    };
  }

  function snapshotMessageRecord(message) {
    const { element, ...snapshot } = message;
    return {
      ...snapshot
    };
  }

  function buildMessageRecordsFromRows(rows, strategy) {
    return rows.reduce((records, row, index) => {
      try {
        const record = buildMessageRecord(row, index, strategy);
        if (record) {
          records.push(record);
        }
      } catch (error) {
        log("Skipping an unsupported Teams message row during refresh.", {
          index,
          messageId: row.getAttribute("data-message-id") || row.id || row.querySelector('[id^="content-"]')?.id || null,
          error: error instanceof Error ? error.message : String(error)
        });
      }

      return records;
    }, []);
  }

  function getVisibleMessageRecords(strategy) {
    if (!strategy) {
      return [];
    }

    const rows = Array.from(document.querySelectorAll(strategy.rowSelector)).filter(isLikelyMessageRow);
    return buildMessageRecordsFromRows(rows, strategy);
  }

  function syncMessageDecorations(message) {
    message.element.classList.add(MESSAGE_CLASS);

    let wrapper = message.element.querySelector(`:scope > .${CHECKBOX_CLASS}`);
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

    message.element.classList.toggle(SELECTED_CLASS, state.active && state.selectedIds.has(message.id));

    const input = wrapper.querySelector("input");
    input.checked = state.selectedIds.has(message.id);
  }

  function handleMessageMouseDown(event) {
    if (!state.active) {
      return;
    }

    const messageElement = event.currentTarget;
    if (isToolbarNode(event.target) || isInteractiveNode(event.target)) {
      return;
    }

    const message = state.messageMap.get(messageElement.dataset.tsmMessageId);
    if (!message) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    clearDomSelection();
    state.lastPointerSelection = {
      id: message.id,
      at: Date.now()
    };
    toggleSelection(message.id, { shiftKey: event.shiftKey });
  }

  function handleMessageClick(event) {
    if (!state.active) {
      return;
    }

    const messageElement = event.currentTarget;
    if (isToolbarNode(event.target) || isInteractiveNode(event.target)) {
      return;
    }

    const message = state.messageMap.get(messageElement.dataset.tsmMessageId);
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

  function bindMessage(message) {
    message.element.dataset.tsmMessageId = message.id;

    if (!message.element.dataset.tsmBound) {
      message.element.dataset.tsmBound = "true";
      message.element.addEventListener("mousedown", handleMessageMouseDown, true);
      message.element.addEventListener("click", handleMessageClick, true);
    }

    syncMessageDecorations(message);
  }

  function refreshMessages() {
    const nextConversationKey = getConversationKey();
    const conversationChanged = Boolean(state.conversationKey) && nextConversationKey !== state.conversationKey;
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

    messages.forEach((message) => {
      state.messageMap.set(message.id, message);
      bindMessage(message);
    });

    cleanupOrphanedDecorations();
    syncDockPosition();
    updateToolbar();
    return messages;
  }

  function hasDisconnectedMessages() {
    return state.messages.length > 0 && state.messages.every((message) => !message.element?.isConnected);
  }

  function checkConversationState() {
    if (state.busy || !state.toolbar) {
      return;
    }

    const nextConversationKey = getConversationKey();
    const conversationChanged = Boolean(state.conversationKey) && nextConversationKey !== state.conversationKey;

    if (conversationChanged || hasDisconnectedMessages()) {
      refreshMessages();
    }
  }

  function cleanupOrphanedDecorations() {
    document.querySelectorAll(`.${MESSAGE_CLASS}`).forEach((element) => {
      if (!state.messages.some((message) => message.element === element)) {
        element.classList.remove(MESSAGE_CLASS, SELECTED_CLASS);
        element.querySelectorAll(`:scope > .${CHECKBOX_CLASS}`).forEach((node) => node.remove());
        element.removeEventListener("mousedown", handleMessageMouseDown, true);
        element.removeEventListener("click", handleMessageClick, true);
        delete element.dataset.tsmMessageId;
        delete element.dataset.tsmBound;
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

  function getMessageRange(anchorId, targetId) {
    const ids = state.messages.map((message) => message.id);
    const anchorIndex = ids.indexOf(anchorId);
    const targetIndex = ids.indexOf(targetId);

    if (anchorIndex === -1 || targetIndex === -1) {
      return [targetId];
    }

    const start = Math.min(anchorIndex, targetIndex);
    const end = Math.max(anchorIndex, targetIndex);
    return ids.slice(start, end + 1);
  }

  function toggleSelection(messageId, options = {}) {
    const currentlySelected = state.selectedIds.has(messageId);
    const targetValue = options.explicitValue ?? !currentlySelected;

    if (options.shiftKey && state.anchorId) {
      const range = getMessageRange(state.anchorId, messageId);
      range.forEach((id) => {
        if (targetValue) {
          state.selectedIds.add(id);
        } else {
          state.selectedIds.delete(id);
        }
      });
    } else if (targetValue) {
      state.selectedIds.add(messageId);
    } else {
      state.selectedIds.delete(messageId);
    }

    state.anchorId = messageId;
    updateSelectionUi();
  }

  function clearSelection() {
    resetSelectionState();
    updateSelectionUi();
  }

  function resetSelectionState() {
    state.selectedIds.clear();
    state.anchorId = null;
    state.lastPointerSelection = null;
    clearDomSelection();
  }

  function setActive(active) {
    state.active = active;
    document.body.classList.toggle(BODY_ACTIVE_CLASS, active);
    clearDomSelection();
    if (active && !state.messages.length) {
      refreshMessages();
    }
    updateToolbar();
    updateSelectionUi();
  }

  function setPanelOpen(open) {
    state.panelOpen = open;
    if (open) {
      syncDockPosition();
      if (!state.messages.length) {
        refreshMessages();
      }
    }
    updateToolbar();
  }

  function updateSelectionUi() {
    state.messages.forEach(syncMessageDecorations);
    updateToolbar();
  }

  function updateToolbar() {
    if (!state.toolbar) {
      return;
    }

    syncDockPosition();
    const selectedCount = state.selectedIds.size;
    const busy = state.busy;
    state.toolbar.classList.toggle(ACTIVE_CLASS, state.active);
    state.toolbar.classList.toggle(PANEL_OPEN_CLASS, state.panelOpen);
    state.toolbar.querySelector('[data-role="count"]').textContent = `${selectedCount} selected`;
    state.toolbar.querySelector('[data-role="active-toggle"]').checked = state.active;
    state.toolbar.querySelector('[data-role="active-toggle"]').disabled = busy;
    state.toolbar.querySelector('[data-role="progress"]').hidden = !busy;
    state.toolbar.querySelector('[data-role="progress"]').textContent = state.busyText;
    state.toolbar.querySelector('[data-action="export-html"]').disabled = busy || selectedCount === 0;
    state.toolbar.querySelector('[data-action="export-md"]').disabled = busy || selectedCount === 0;
    state.toolbar.querySelector('[data-action="export-full-md"]').disabled = busy;
    state.toolbar.querySelector('[data-action="export-full-html"]').disabled = busy;
    state.toolbar.querySelector('[data-action="copy-md"]').disabled = busy || !state.active || selectedCount === 0;
    state.toolbar.querySelector('[data-action="clear"]').disabled = busy || selectedCount === 0;
    if (state.launcher) {
      state.launcher.textContent = "Export";
      state.launcher.setAttribute("aria-expanded", state.panelOpen ? "true" : "false");
      state.launcher.disabled = busy;
    }
    if (state.quickCopyButton) {
      state.quickCopyButton.hidden = !(selectedCount > 0 && !state.panelOpen);
      state.quickCopyButton.textContent = selectedCount > 0 ? `Copy ${selectedCount}` : "Copy MD";
      state.quickCopyButton.disabled = busy || selectedCount === 0;
    }
  }

  function getSelectedMessages() {
    return state.messages.filter((message) => state.selectedIds.has(message.id));
  }

  function getExportMessageCountLabel(count, scope) {
    if (scope === "full-chat") {
      return `Full chat history messages: ${count}`;
    }

    return `Selected messages: ${count}`;
  }

  function getExportMessageSummary(count, scope) {
    if (scope === "full-chat") {
      return `${count} message${count === 1 ? "" : "s"} captured from the full chat history.`;
    }

    return `${count} message${count === 1 ? "" : "s"} selected.`;
  }

  function getExportScopeSuffix(scope) {
    return scope === "full-chat" ? "-full-chat" : "";
  }

  function renderHtmlDocument(messages, meta = {}) {
    const title = getConversationTitle();
    const exportedAt = formatLocalDateTime(new Date());
    const scope = meta.scope || "selection";
    const articles = messages
      .map((message) => {
        const safeAuthor = escapeHtml(message.author);
        const safeTime = escapeHtml(message.timeLabel || message.dateTime || "");
        const quoteHtml = renderQuotedReplyHtml(message.quote);
        const reactionsHtml = renderReactionsHtml(message.reactions);
        return `
          <article class="message">
            <header>
              <strong>${safeAuthor}</strong>
              <time datetime="${escapeHtml(message.dateTime || "")}">${safeTime}</time>
            </header>
            <section class="body">${quoteHtml}${message.html || `<p>${escapeHtml(message.plainText)}</p>`}${reactionsHtml}</section>
          </article>
        `;
      })
      .join("\n");

    return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(title)} - Teams export</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body {
        margin: 0;
        padding: 32px;
        background: #f5f7fb;
        color: #132238;
        font: 15px/1.55 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      main {
        max-width: 920px;
        margin: 0 auto;
        display: grid;
        gap: 16px;
      }
      .summary {
        padding: 20px 24px;
        border-radius: 18px;
        background: #fff;
        box-shadow: 0 10px 30px rgba(16, 35, 63, 0.08);
      }
      .message {
        padding: 18px 20px;
        border-radius: 18px;
        background: #fff;
        box-shadow: 0 10px 30px rgba(16, 35, 63, 0.08);
      }
      .message header {
        display: flex;
        justify-content: space-between;
        gap: 16px;
        margin-bottom: 12px;
      }
      .message time {
        color: #5a6980;
        white-space: nowrap;
      }
      .body pre {
        overflow: auto;
        padding: 12px;
        border-radius: 12px;
        background: #eef3fb;
      }
      .body blockquote {
        margin-left: 0;
        margin-bottom: 14px;
        padding-left: 14px;
        border-left: 3px solid #0d61d8;
        color: #314865;
        background: rgba(13, 97, 216, 0.04);
        padding-top: 10px;
        padding-bottom: 10px;
        border-radius: 0 12px 12px 0;
      }
      .reply-meta {
        margin: 0 0 8px;
        color: #0d457f;
        font-size: 13px;
      }
      [data-tsm-placeholder="image"] {
        display: inline-block;
        padding: 4px 8px;
        border-radius: 999px;
        background: rgba(19, 34, 56, 0.08);
        color: #49607f;
        font-size: 12px;
        font-style: italic;
      }
      .reactions {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 14px;
      }
      .reaction {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 10px;
        border-radius: 999px;
        background: rgba(19, 34, 56, 0.06);
        color: #1b3455;
        font-size: 12px;
      }
    </style>
  </head>
  <body>
    <main>
      <section class="summary">
        <h1>${escapeHtml(title)}</h1>
        <p>Exported from Microsoft Teams on ${escapeHtml(exportedAt)}.</p>
        <p>${escapeHtml(getExportMessageSummary(messages.length, scope))}</p>
        <p>Source: <a href="${escapeHtml(location.href)}">${escapeHtml(location.href)}</a></p>
      </section>
      ${articles}
    </main>
  </body>
</html>`;
  }

  function inlineMarkdown(text, activeHref) {
    return text.replace(/\s+/g, " ").trim() || activeHref || "";
  }

  function formatQuotedReplyLabel(quote) {
    const parts = [];
    if (quote.author) {
      parts.push(quote.author);
    }
    if (quote.timeLabel) {
      parts.push(quote.timeLabel);
    }

    return parts.length ? `Replying to ${parts.join(" | ")}` : "Replying to";
  }

  function renderQuotedReplyHtml(quote) {
    if (!quote?.text) {
      return "";
    }

    const safeLabel = escapeHtml(formatQuotedReplyLabel(quote));
    const safeText = escapeHtml(quote.text);

    return `
      <blockquote>
        <p class="reply-meta"><strong>${safeLabel}</strong></p>
        <p>${safeText}</p>
      </blockquote>
    `;
  }

  function renderReactionsHtml(reactions) {
    if (!reactions?.length) {
      return "";
    }

    const chips = reactions
      .map((reaction) => {
        const summary = reaction.actors.length
          ? `${reaction.emoji} ${escapeHtml(reaction.name)} by ${escapeHtml(reaction.actors.join(", "))}`
          : `${reaction.emoji} ${escapeHtml(reaction.name)} x${reaction.count}`;

        return `<span class="reaction">${summary.trim()}</span>`;
      })
      .join("");

    return `<div class="reactions">${chips}</div>`;
  }

  function renderQuotedReplyMarkdown(quote) {
    if (!quote?.text) {
      return "";
    }

    const lines = [`> ${formatQuotedReplyLabel(quote)}`];

    quote.text
      .split("\n")
      .map((line) => normalizeText(line))
      .filter(Boolean)
      .forEach((line) => {
        lines.push(`> ${line}`);
      });

    return lines.join("\n");
  }

  function renderReactionsMarkdown(reactions) {
    if (!reactions?.length) {
      return "";
    }

    const parts = reactions.map((reaction) => {
      if (reaction.actors.length) {
        return `${reaction.emoji} ${reaction.name} by ${reaction.actors.join(", ")}`;
      }

      return `${reaction.emoji} ${reaction.name} x${reaction.count}`;
    });

    return `Reactions: ${parts.join("; ")}`;
  }

  function nodeToMarkdown(node, context = {}) {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent || "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return "";
    }

    const element = node;
    const tag = element.tagName.toLowerCase();
    const children = Array.from(element.childNodes).map((child) => nodeToMarkdown(child, context)).join("");

    if (isMentionElement(element)) {
      return formatMentionLabel(element);
    }

    if (tag === "img") {
      if (isEmojiImage(element)) {
        return element.getAttribute("alt") || "";
      }

      if (isDecorativeImage(element)) {
        return "";
      }

      return "[Image omitted]";
    }

    if (tag === "br") {
      return "  \n";
    }

    if (tag === "strong" || tag === "b") {
      return `**${children.trim()}**`;
    }

    if (tag === "em" || tag === "i") {
      return `*${children.trim()}*`;
    }

    if (tag === "code" && element.parentElement?.tagName.toLowerCase() !== "pre") {
      return `\`${children.trim()}\``;
    }

    if (tag === "pre") {
      return `\n\`\`\`\n${element.textContent?.trim() || ""}\n\`\`\`\n`;
    }

    if (tag === "a") {
      const text = inlineMarkdown(children, element.href);
      const href = element.href || "";
      return href ? `[${text}](${href})` : text;
    }

    if (tag === "li") {
      return `- ${children.trim()}\n`;
    }

    if (tag === "ul" || tag === "ol") {
      return `\n${children.trimEnd()}\n`;
    }

    if (tag === "blockquote") {
      return `\n${children
        .split("\n")
        .filter(Boolean)
        .map((line) => `> ${line}`)
        .join("\n")}\n`;
    }

    if (tag === "p" || tag === "div" || tag === "section") {
      if (tag === "div" && (isMentionElement(element) || element.dataset.tsmMention === "true")) {
        return children.trim();
      }

      return `${children.trim()}\n\n`;
    }

    return children;
  }

  function elementToMarkdown(element, strategy, fallbackText) {
    if (!element) {
      return fallbackText;
    }

    const clone = getPreparedContentClone(element, strategy);
    const markdown = Array.from(clone.childNodes)
      .map((node) => nodeToMarkdown(node))
      .join("")
      .replace(/\n{3,}/g, "\n\n")
      .trim();

    return markdown || fallbackText;
  }

  function renderMarkdown(messages, meta = {}) {
    const title = getConversationTitle();
    const scope = meta.scope || "selection";
    const lines = [
      `# ${title}`,
      "",
      `- Exported from Microsoft Teams on ${formatLocalDateTime(new Date())}`,
      `- Source URL: ${location.href}`,
      `- ${getExportMessageCountLabel(messages.length, scope)}`,
      ""
    ];

    messages.forEach((message) => {
      const headingBits = [message.author];
      if (message.timeLabel || message.dateTime) {
        headingBits.push(message.timeLabel || message.dateTime);
      }

      lines.push(`## ${headingBits.join(" | ")}`);
      lines.push("");
      if (message.quote?.text) {
        lines.push(renderQuotedReplyMarkdown(message.quote));
        lines.push("");
      }
      lines.push(message.markdown || elementToMarkdown(message.element, state.strategy, message.plainText));
      if (message.reactions?.length) {
        lines.push("");
        lines.push(renderReactionsMarkdown(message.reactions));
      }
      lines.push("");
    });

    return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }

  function triggerDownload(filename, mimeType, content) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  async function writeClipboardText(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (error) {
      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.setAttribute("readonly", "true");
      textarea.style.position = "fixed";
      textarea.style.top = "0";
      textarea.style.left = "0";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();

      const copied = document.execCommand("copy");
      textarea.remove();

      if (!copied) {
        throw error;
      }

      return true;
    }
  }

  function createExportPayload(format, messages, meta = {}) {
    const scope = meta.scope || "selection";
    const titlePart = filenameSafe(getConversationTitle()) || "teams-conversation";
    const timestampPart = new Date().toISOString().replace(/[:.]/g, "-");
    const scopeSuffix = getExportScopeSuffix(scope);

    if (format === "html") {
      return {
        format,
        filename: `${titlePart}${scopeSuffix}-${timestampPart}.html`,
        content: renderHtmlDocument(messages, meta),
        count: messages.length,
        scope
      };
    }

    return {
      format: "md",
      filename: `${titlePart}${scopeSuffix}-${timestampPart}.md`,
      content: renderMarkdown(messages, meta),
      count: messages.length,
      scope
    };
  }

  function commitExportPayload(payload, options = {}) {
    if (options.download !== false) {
      triggerDownload(
        payload.filename,
        payload.format === "html" ? "text/html;charset=utf-8" : "text/markdown;charset=utf-8",
        payload.content
      );
    }

    state.lastExport = payload;
    return payload;
  }

  function exportSelection(format) {
    const messages = getSelectedMessages();
    if (!messages.length) {
      log("No selected messages to export.");
      return null;
    }

    return commitExportPayload(createExportPayload(format, messages, { scope: "selection" }));
  }

  async function copyMarkdown() {
    const messages = getSelectedMessages();
    if (!messages.length) {
      return false;
    }

    const markdown = renderMarkdown(messages, { scope: "selection" });
    await writeClipboardText(markdown);
    state.lastExport = { format: "clipboard-md", content: markdown, count: messages.length };
    setActive(false);
    setPanelOpen(false);
    return true;
  }

  function setBusy(active, text = "") {
    state.busy = active;
    state.busyText = active ? text : "";
    updateToolbar();
  }

  function waitForDelay(ms) {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
  }

  function getRunwayElement() {
    return document.querySelector(RUNWAY_SELECTOR);
  }

  function isElementVisible(element) {
    if (!element || !element.isConnected) {
      return false;
    }

    const styles = window.getComputedStyle(element);
    if (styles.display === "none" || styles.visibility === "hidden" || styles.opacity === "0") {
      return false;
    }

    if (element.hidden || element.getAttribute("aria-hidden") === "true") {
      return false;
    }

    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function isScrollableElement(element) {
    if (!element || element === document.body) {
      return false;
    }

    const styles = window.getComputedStyle(element);
    return /(auto|scroll|overlay)/.test(styles.overflowY) && element.scrollHeight > element.clientHeight + 8;
  }

  function getMessageScrollContainer() {
    const runway = getRunwayElement();
    let current = runway;

    while (current) {
      if (isScrollableElement(current)) {
        return current;
      }

      current = current.parentElement;
    }

    return document.scrollingElement || document.documentElement;
  }

  function toComparableTime(value) {
    const numericTime = Date.parse(value || "");
    return Number.isFinite(numericTime) ? numericTime : null;
  }

  function toComparableNumericId(value) {
    if (!/^\d+$/.test(String(value || ""))) {
      return null;
    }

    const numericId = Number(value);
    return Number.isFinite(numericId) ? numericId : null;
  }

  function orderMessagesForExport(messages) {
    return [...messages].sort((left, right) => {
      const leftTime = toComparableTime(left.dateTime);
      const rightTime = toComparableTime(right.dateTime);
      if (leftTime !== null && rightTime !== null && leftTime !== rightTime) {
        return leftTime - rightTime;
      }

      const leftId = toComparableNumericId(left.id);
      const rightId = toComparableNumericId(right.id);
      if (leftId !== null && rightId !== null && leftId !== rightId) {
        return leftId - rightId;
      }

      const leftFallback = left.captureOrder ?? left.index ?? 0;
      const rightFallback = right.captureOrder ?? right.index ?? 0;
      return rightFallback - leftFallback;
    });
  }

  function collectVisibleMessagesIntoMap(strategy, harvestedMap, captureOrderStart = 0) {
    const visibleMessages = getVisibleMessageRecords(strategy);
    let captureOrder = captureOrderStart;

    visibleMessages.forEach((message) => {
      if (!harvestedMap.has(message.id)) {
        harvestedMap.set(message.id, {
          ...snapshotMessageRecord(message),
          captureOrder
        });
        captureOrder += 1;
      }
    });

    return {
      visibleMessages,
      nextCaptureOrder: captureOrder
    };
  }

  function getVisibleHistoryRows(strategy) {
    if (!strategy) {
      return [];
    }

    return Array.from(document.querySelectorAll(strategy.rowSelector)).filter(isLikelyMessageRow);
  }

  function getHistorySignature(messages, scrollContainer) {
    const firstId = messages[0]?.id || "none";
    const lastId = messages[messages.length - 1]?.id || "none";
    const top = Math.round(scrollContainer.scrollTop || 0);
    const height = Math.round(scrollContainer.scrollHeight || 0);
    return `${firstId}|${lastId}|${messages.length}|${top}|${height}`;
  }

  function getRowHistorySignature(rows, scrollContainer) {
    const firstId = rows[0] ? getMessageId(rows[0], 0) : "none";
    const lastId = rows.length ? getMessageId(rows[rows.length - 1], rows.length - 1) : "none";
    const top = Math.round(scrollContainer.scrollTop || 0);
    const height = Math.round(scrollContainer.scrollHeight || 0);
    return `${firstId}|${lastId}|${rows.length}|${top}|${height}`;
  }

  function hasActiveHistoryLoadingIndicator(strategy, scrollContainer) {
    const roots = [getRunwayElement(), scrollContainer, scrollContainer?.parentElement, scrollContainer?.closest("[role='main']")];
    const rowSelector = strategy?.rowSelector || "";
    const seen = new Set();

    return roots
      .filter(Boolean)
      .some((root) => {
        if (seen.has(root)) {
          return false;
        }

        seen.add(root);

        return Array.from(root.querySelectorAll(HISTORY_LOADING_SELECTOR)).some((candidate) => {
          if (!isElementVisible(candidate)) {
            return false;
          }

          if (rowSelector && candidate.closest(rowSelector)) {
            return false;
          }

          return true;
        });
      });
  }

  async function waitForHistoryToSettle(strategy, scrollContainer) {
    const nearTop = scrollContainer.scrollTop <= 4;
    const quietWindowMs = nearTop ? HISTORY_SETTLE_TOP_MS : HISTORY_SETTLE_MS;
    const maxWaitMs = nearTop ? HISTORY_MAX_WAIT_TOP_MS : HISTORY_MAX_WAIT_MS;
    const startedAt = performance.now();
    let lastChangeAt = startedAt;
    let lastSignature = getRowHistorySignature(getVisibleHistoryRows(strategy), scrollContainer);
    let lastLoadingState = hasActiveHistoryLoadingIndicator(strategy, scrollContainer);

    while (performance.now() - startedAt < maxWaitMs) {
      await waitForDelay(HISTORY_POLL_INTERVAL_MS);

      const currentSignature = getRowHistorySignature(getVisibleHistoryRows(strategy), scrollContainer);
      const loadingIndicatorVisible = hasActiveHistoryLoadingIndicator(strategy, scrollContainer);
      const now = performance.now();

      if (currentSignature !== lastSignature || loadingIndicatorVisible !== lastLoadingState) {
        lastChangeAt = now;
        lastSignature = currentSignature;
        lastLoadingState = loadingIndicatorVisible;
      }

      if (!loadingIndicatorVisible && now - lastChangeAt >= quietWindowMs) {
        break;
      }
    }

    return getRowHistorySignature(getVisibleHistoryRows(strategy), scrollContainer);
  }

  async function harvestFullChatMessages() {
    const strategy = state.strategy || selectStrategy().strategy;
    if (!strategy) {
      return [];
    }

    const scrollContainer = getMessageScrollContainer();
    const originalDistanceFromBottom = scrollContainer.scrollHeight - scrollContainer.scrollTop;
    const harvestedMap = new Map();
    let captureOrder = 0;
    let stagnantPasses = 0;
    let previousSignature = "";

    try {
      for (let pass = 0; pass < 160; pass += 1) {
        const beforeCollect = collectVisibleMessagesIntoMap(strategy, harvestedMap, captureOrder);
        captureOrder = beforeCollect.nextCaptureOrder;
        const sizeBeforeScroll = harvestedMap.size;
        setBusy(true, `Loading full chat history... ${harvestedMap.size} messages captured`);

        const beforeTop = scrollContainer.scrollTop;
        const scrollStep = Math.max(260, Math.round((scrollContainer.clientHeight || window.innerHeight || 800) * 0.82));
        scrollContainer.scrollTop = Math.max(0, beforeTop - scrollStep);
        await waitForHistoryToSettle(strategy, scrollContainer);

        const afterCollect = collectVisibleMessagesIntoMap(strategy, harvestedMap, captureOrder);
        captureOrder = afterCollect.nextCaptureOrder;
        setBusy(true, `Loading full chat history... ${harvestedMap.size} messages captured`);

        const currentSignature = getHistorySignature(afterCollect.visibleMessages, scrollContainer);
        const reachedTop = scrollContainer.scrollTop <= 4;
        const discoveredNewData = harvestedMap.size > sizeBeforeScroll || currentSignature !== previousSignature;

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
      scrollContainer.scrollTop = Math.max(0, scrollContainer.scrollHeight - originalDistanceFromBottom);
      await waitForDelay(280);
      refreshMessages();
    }
  }

  async function exportFullHistory(format = "md", options = {}) {
    if (state.busy) {
      return null;
    }

    const runway = getRunwayElement();
    if (!runway) {
      log("No Teams message runway found for full-history export.");
      return null;
    }

    setBusy(true, "Loading full chat history...");

    try {
      const messages = await harvestFullChatMessages();
      if (!messages.length) {
        return null;
      }

      const payload = commitExportPayload(createExportPayload(format, messages, { scope: "full-chat" }), options);
      if (options.closePanel !== false) {
        setPanelOpen(false);
      }

      return {
        ...payload,
        messages
      };
    } finally {
      setBusy(false);
    }
  }

  function handleSelectStart(event) {
    if (!state.active) {
      return;
    }

    if (isToolbarNode(event.target) || !isWithinMessageRunway(event.target)) {
      return;
    }

    event.preventDefault();
  }

  function ensureGlobalHandlers() {
    if (!state.selectStartHandler) {
      state.selectStartHandler = handleSelectStart;
      document.addEventListener("selectstart", state.selectStartHandler, true);
    }

    if (!state.resizeHandler) {
      state.resizeHandler = () => {
        syncDockPosition();
        applyTheme();
      };
      window.addEventListener("resize", state.resizeHandler, { passive: true });
    }

    if (!state.themeMediaQuery && window.matchMedia) {
      state.themeMediaQuery = window.matchMedia("(prefers-color-scheme: dark)");
      state.themeMediaQuery.addEventListener?.("change", applyTheme);
    }

    if (!state.focusHandler) {
      state.focusHandler = checkConversationState;
      window.addEventListener("focus", state.focusHandler, { passive: true });
    }

    if (!state.visibilityHandler) {
      state.visibilityHandler = () => {
        if (document.visibilityState !== "hidden") {
          checkConversationState();
        }
      };
      document.addEventListener("visibilitychange", state.visibilityHandler, { passive: true });
    }

    if (!state.pageShowHandler) {
      state.pageShowHandler = checkConversationState;
      window.addEventListener("pageshow", state.pageShowHandler, { passive: true });
    }

    if (!state.popStateHandler) {
      state.popStateHandler = checkConversationState;
      window.addEventListener("popstate", state.popStateHandler, { passive: true });
    }

    if (!state.hashChangeHandler) {
      state.hashChangeHandler = checkConversationState;
      window.addEventListener("hashchange", state.hashChangeHandler, { passive: true });
    }

    if (!state.conversationPollTimer) {
      state.conversationPollTimer = window.setInterval(() => {
        if (document.visibilityState !== "hidden") {
          checkConversationState();
        }
      }, CONVERSATION_POLL_MS);
    }
  }

  function setupExtensionBridge() {
    if (!globalThis.chrome?.runtime?.onMessage || window.__teamsMessageExporterBridgeBound) {
      return;
    }

    chrome.runtime.onMessage.addListener((message) => {
      if (!message || message.type !== "teams-export/toggle-panel") {
        return;
      }

      setPanelOpen(!state.panelOpen);
    });

    window.__teamsMessageExporterBridgeBound = true;
  }

  function observeDom() {
    const observer = new MutationObserver((mutations) => {
      const shouldRefresh = mutations.some(
        (mutation) => mutation.type === "childList" || mutation.type === "characterData"
      );
      const shouldApplyTheme = mutations.some(
        (mutation) =>
          mutation.type === "attributes" &&
          (mutation.target === document.documentElement || mutation.target === document.body)
      );
      const shouldResyncDock = mutations.some(
        (mutation) =>
          mutation.type === "childList" ||
          mutation.target === document.documentElement ||
          mutation.target === document.body
      );

      if (shouldRefresh) {
        debounceRefresh();
      }

      if (shouldApplyTheme) {
        applyTheme();
      }

      if (shouldResyncDock) {
        syncDockPosition();
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

  function destroy() {
    state.observer?.disconnect();
    window.clearTimeout(state.refreshTimer);
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
    if (state.themeMediaQuery) {
      state.themeMediaQuery.removeEventListener?.("change", applyTheme);
      state.themeMediaQuery = null;
    }
    document.body?.classList.remove(BODY_ACTIVE_CLASS);
    state.styleElement?.remove();
    state.dock?.remove();

    document.querySelectorAll(`.${MESSAGE_CLASS}`).forEach((element) => {
      element.classList.remove(MESSAGE_CLASS, SELECTED_CLASS);
      element.querySelectorAll(`:scope > .${CHECKBOX_CLASS}`).forEach((node) => node.remove());
      element.removeEventListener("mousedown", handleMessageMouseDown, true);
      element.removeEventListener("click", handleMessageClick, true);
      delete element.dataset.tsmMessageId;
      delete element.dataset.tsmBound;
    });

    delete window[INSTANCE_KEY];
  }

  function registerApi() {
    window[INSTANCE_KEY] = {
      state,
      refresh: refreshMessages,
      setActive,
      setPanelOpen,
      togglePanel: () => setPanelOpen(!state.panelOpen),
      clearSelection,
      exportSelection,
      exportFullHistory,
      renderHtml: () => renderHtmlDocument(getSelectedMessages()),
      renderMarkdown: () => renderMarkdown(getSelectedMessages()),
      getSelectedMessages: () =>
        getSelectedMessages().map(({ element, ...message }) => ({
          ...message
        })),
      destroy
    };
  }

  function start() {
    if (!document.head || !document.body) {
      return;
    }

    if (!state.toolbar) {
      ensureStyles();
      createToolbar();
      ensureGlobalHandlers();
      setupExtensionBridge();
      observeDom();
    }

    applyTheme();
    refreshMessages();
    registerApi();
    log("Prototype loaded.");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }
})();
