export interface Strategy {
  name: string;
  rowSelector: string;
  contentSelector: string;
  authorSelector: string;
  timeSelector: string;
}

export interface WorkerCapturedEmotionEntry {
  key: string | null;
  userIds: string[];
}

export interface WorkerCapturedEmotionSummaryEntry {
  key: string | null;
  count: number | null;
}

/** A chat message captured from the worker intercept bridge. */
export interface WorkerCapturedMessage {
  id: string;
  clientMessageId: string | null;
  originalArrivalTime: string | null;
  author: string | null;
  authorId: string | null;
  messageType: string | null;
  content: string | null;
  subject: string | null;
  threadType: string | null;
  isDeleted: boolean;
  editedTime: string | null;
  emotionsSummary: WorkerCapturedEmotionSummaryEntry[] | null;
  emotions: WorkerCapturedEmotionEntry[] | null;
  diverseEmotions: WorkerCapturedEmotionEntry[] | null;
  quotedMessages: unknown | null;
}

export interface TimeMeta {
  label: string;
  dateTime: string;
}

export interface QuotedReply {
  author: string;
  timeLabel: string;
  text: string;
}

export interface ReactionInfo {
  emoji: string;
  name: string;
  count: number;
  actors: string[];
}

export interface MessageRecord {
  id: string;
  index: number;
  element: HTMLElement;
  author: string;
  timeLabel: string;
  dateTime: string;
  subject: string;
  quote: QuotedReply | null;
  reactions: ReactionInfo[];
  html: string;
  markdown: string;
  plainText: string;
}

export interface MessageSnapshot {
  id: string;
  index: number;
  author: string;
  timeLabel: string;
  dateTime: string;
  subject: string;
  quote: QuotedReply | null;
  reactions: ReactionInfo[];
  html: string;
  markdown: string;
  plainText: string;
  captureOrder?: number;
}

export interface ExportPayload {
  format: string;
  filename: string;
  content: string;
  count: number;
  scope: string;
}

export interface ExportFullHistoryResult extends ExportPayload {
  messages: MessageSnapshot[];
}

export interface ToggleSelectionOptions {
  shiftKey?: boolean;
  explicitValue?: boolean;
}

export interface ExportMeta {
  scope?: string;
  title?: string;
  sourceUrl?: string;
}

export interface RgbColor {
  red: number;
  green: number;
  blue: number;
  alpha: number;
}

export interface ExporterState {
  active: boolean;
  panelOpen: boolean;
  dock: HTMLElement | null;
  toolbar: HTMLElement | null;
  launcher: HTMLButtonElement | null;
  quickCopyButton: HTMLButtonElement | null;
  styleElement: HTMLStyleElement | null;
  observer: MutationObserver | null;
  themeMediaQuery: MediaQueryList | null;
  refreshTimer: number | null;
  conversationPollTimer: number | null;
  startupTimers: number[];
  selectStartHandler: ((event: Event) => void) | null;
  resizeHandler: (() => void) | null;
  focusHandler: (() => void) | null;
  visibilityHandler: (() => void) | null;
  pageShowHandler: (() => void) | null;
  popStateHandler: (() => void) | null;
  hashChangeHandler: (() => void) | null;
  busy: boolean;
  busyText: string;
  messages: MessageRecord[];
  messageMap: Map<string, MessageRecord>;
  selectedIds: Set<string>;
  anchorId: string | null;
  strategy: Strategy | null;
  theme: string;
  conversationKey: string;
  lastExport: ExportPayload | null;
  lastPointerSelection: { id: string; at: number } | null;
}

export interface ExporterCallbacks {
  updateToolbar(): void;
  refreshMessages(): MessageRecord[];
  syncDockPosition(): void;
  applyTheme(): void;
  setBusy(active: boolean, text?: string): void;
  checkConversationState(): void;
  scheduleStartupStabilization(): void;
  debounceRefresh(): void;
  setActive(active: boolean): void;
  setPanelOpen(open: boolean): void;
  clearSelection(): void;
  exportSelection(format: string): ExportPayload | null;
  exportFullHistory(format: string): Promise<ExportFullHistoryResult | null>;
  copyMarkdown(): Promise<boolean>;
  toggleSelection(messageId: string, options?: ToggleSelectionOptions): void;
}

export interface ToolbarActions {
  onClear(): void;
  onExportHtml(): void;
  onExportMarkdown(): void;
  onExportFullMarkdown(): void;
  onExportFullHtml(): void;
  onCopyMarkdown(): void;
  onTogglePanel(): void;
  onToggleActive(active: boolean): void;
}

export interface ToolbarElements {
  dock: HTMLElement;
  toolbar: HTMLElement;
  launcher: HTMLButtonElement;
  quickCopyButton: HTMLButtonElement;
}
