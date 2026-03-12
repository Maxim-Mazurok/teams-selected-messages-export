import type { Strategy } from "./types.js";

export const INSTANCE_KEY = "__teamsMessageExporter";

export const UI_CLASS = "tsm-export";
export const DOCK_CLASS = `${UI_CLASS}__dock`;
export const PANEL_CLASS = `${UI_CLASS}__panel`;
export const PANEL_OPEN_CLASS = `${PANEL_CLASS}--open`;
export const LAUNCHER_ROW_CLASS = `${UI_CLASS}__launcher-row`;
export const LAUNCHER_CLASS = `${UI_CLASS}__launcher`;
export const QUICK_COPY_CLASS = `${UI_CLASS}__quick-copy`;
export const HEADER_CLASS = `${UI_CLASS}__header`;
export const STATUS_ROW_CLASS = `${UI_CLASS}__status-row`;
export const PROGRESS_CLASS = `${UI_CLASS}__progress`;
export const TOGGLE_CLASS = `${UI_CLASS}__toggle`;
export const COUNT_CLASS = `${UI_CLASS}__count`;
export const ACTION_GRID_CLASS = `${UI_CLASS}__action-grid`;
export const MESSAGE_CLASS = `${UI_CLASS}__message`;
export const SELECTED_CLASS = `${MESSAGE_CLASS}--selected`;
export const ACTIVE_CLASS = `${PANEL_CLASS}--active`;
export const CHECKBOX_CLASS = `${UI_CLASS}__checkbox`;
export const HIDDEN_CHECKBOX_CLASS = `${CHECKBOX_CLASS}--hidden`;
export const BODY_ACTIVE_CLASS = `${UI_CLASS}--selection-active`;

export const RUNWAY_SELECTOR = [
  '[data-tid="message-pane-list-runway"]',
  '#chat-pane-list',
  '[data-tid="channel-pane-runway"]',
  '[data-tid="channel-replies-runway"]',
  '[data-tid="reply-thread-pane-runway"]',
  '[data-tid="reply-chain-pane-runway"]',
  '[data-tid="thread-pane-list-runway"]',
  '[data-tid="message-pane-list"]'
].join(", ");

export const HISTORY_LOADING_SELECTOR = [
  '[role="progressbar"]',
  '[aria-busy="true"]',
  '[data-tid*="loading" i]',
  '[data-tid*="spinner" i]',
  '[data-tid*="loader" i]',
  '[data-testid*="loading" i]',
  '[data-testid*="spinner" i]',
  '[data-testid*="loader" i]'
].join(", ");

export const HISTORY_POLL_INTERVAL_MS = 120;
export const HISTORY_SETTLE_MS = 260;
export const HISTORY_SETTLE_TOP_MS = 720;
export const HISTORY_MAX_WAIT_MS = 1_100;
export const HISTORY_MAX_WAIT_TOP_MS = 3_200;
export const HISTORY_TOP_STAGNANT_PASSES = 2;
export const HISTORY_CONTENT_STAGNANT_PASSES = 8;
export const HISTORY_MAX_PASSES = 2_000;

export const HEADER_SLOT_SELECTORS = [
  "#titlebar-end-slot-start",
  '[data-tid="titlebar-end-slot"]'
];

export const FALLBACK_DOCK_TOP = 10;
export const FALLBACK_DOCK_RIGHT = 200;
export const CONVERSATION_POLL_MS = 750;

export const CONTENT_PRUNE_SELECTOR = [
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
  'button[aria-label*="more" i]',
  '[data-tid="url-preview"]',
  '[data-tid="subject-line"]'
].join(", ");

export const STRATEGIES: Strategy[] = [
  {
    name: "teams-data-tid-chat-pane-item",
    rowSelector: '[data-tid="chat-pane-item"]',
    contentSelector:
      '[data-tid="chat-pane-message"], [data-message-content], [data-tid="message-body"], [data-tid="message-content"]',
    authorSelector: '[data-tid="message-author-name"], [data-tid="message-author"]',
    timeSelector: 'time, [data-tid="message-timestamp"], [data-tid="timestamp"]'
  },
  {
    name: "teams-channel-pane-message",
    rowSelector:
      '[data-tid="channel-pane-message"], [data-tid="channel-replies-pane-message"], [data-tid="response-surface"] [role="group"]',
    contentSelector: '[data-tid="message-body"]',
    authorSelector: '[data-tid="post-message-subheader"] span[id^="author-"], [data-tid="reply-message-header"] span[id^="author-"]',
    timeSelector: '[data-tid="timestamp"], time'
  },
  {
    name: "teams-data-tid-message-item",
    rowSelector:
      '[data-tid="message-pane-list-item"], [data-tid="message-item"], [data-tid="chat-message"]',
    contentSelector:
      '[data-tid="message-body"], [data-tid="message-content"], [data-tid="chat-pane-message"]',
    authorSelector: '[data-tid="message-author-name"], [data-tid="message-author"]',
    timeSelector: 'time, [data-tid="message-timestamp"], [data-tid="timestamp"]'
  },
  {
    name: "generic-listitem",
    rowSelector: 'div[role="listitem"], li[role="listitem"], [data-message-id]',
    contentSelector:
      '[data-tid="message-body"], [data-tid="message-content"], [dir="auto"], p, div',
    authorSelector:
      '[data-tid="message-author-name"], [data-tid="message-author"], [aria-label*="sent by" i]',
    timeSelector: 'time, [aria-label*=" at " i], [data-tid*="timestamp"]'
  }
];
