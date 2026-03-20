import type { ExporterState, ExporterCallbacks, MessageRecord, ExportFullHistoryResult } from "./types.js";

export const state: ExporterState = {
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
  messageScanTimer: null,
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
  lastPointerSelection: null,
  exportOptions: { includeLinks: false }
};

export const callbacks: ExporterCallbacks = {
  updateToolbar: () => {},
  refreshMessages: (): MessageRecord[] => [],
  syncDockPosition: () => {},
  applyTheme: () => {},
  setBusy: () => {},
  checkConversationState: () => {},
  scheduleStartupStabilization: () => {},
  debounceRefresh: () => {},
  setActive: () => {},
  setPanelOpen: () => {},
  clearSelection: () => {},
  exportSelection: () => null,
  exportFullHistory: () => Promise.resolve(null as ExportFullHistoryResult | null),
  copyMarkdown: () => Promise.resolve(false),
  toggleSelection: () => {}
};

export function log(...arguments_: unknown[]): void {
  console.log("[teams-export]", ...arguments_);
}
