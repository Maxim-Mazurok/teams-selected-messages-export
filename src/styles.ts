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
  MESSAGE_CLASS,
  SELECTED_CLASS,
  ACTIVE_CLASS,
  CHECKBOX_CLASS,
  HIDDEN_CHECKBOX_CLASS,
  BODY_ACTIVE_CLASS,
  OPTIONS_ROW_CLASS,
  RUNWAY_SELECTOR
} from "./constants.js";

export function generateStylesheet(): string {
  return `
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
      --tsm-panel-highlight-shadow: 0 18px 46px rgba(0, 84, 190, 0.2);
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

    .${OPTIONS_ROW_CLASS} {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 14px;
      padding: 8px 12px;
      border-radius: 14px;
      background: var(--tsm-panel-surface-muted);
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

    ${RUNWAY_SELECTOR.split(",")
      .map((selector) => selector.trim())
      .flatMap((selector) => [
        `body.${BODY_ACTIVE_CLASS} ${selector}`,
        `body.${BODY_ACTIVE_CLASS} ${selector} *`
      ])
      .join(",\n    ")} {
      user-select: none;
      -webkit-user-select: none;
    }
  `;
}
