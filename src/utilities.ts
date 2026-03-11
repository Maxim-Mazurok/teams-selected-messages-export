import type { RgbColor } from "./types.js";
import { DOCK_CLASS, RUNWAY_SELECTOR } from "./constants.js";

export function normalizeText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function formatLocalDateTime(value: Date | string): string {
  const date = value instanceof Date ? value : new Date(value);
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "long",
    timeStyle: "medium"
  }).format(date);
}

export function filenameSafe(value: string): string {
  return value
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
}

export function parseRgbColor(value: string | null | undefined): RgbColor | null {
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

export function relativeLuminanceChannel(channel: number): number {
  const normalized = channel / 255;
  if (normalized <= 0.03928) {
    return normalized / 12.92;
  }

  return ((normalized + 0.055) / 1.055) ** 2.4;
}

export function isDarkRgb(color: RgbColor | null): boolean | null {
  if (!color || color.alpha === 0) {
    return null;
  }

  const luminance =
    0.2126 * relativeLuminanceChannel(color.red) +
    0.7152 * relativeLuminanceChannel(color.green) +
    0.0722 * relativeLuminanceChannel(color.blue);

  return luminance < 0.35;
}

export function queryText(root: Element, selectorList: string): string {
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

export function inlineMarkdown(text: string, activeHref?: string): string {
  return text.replace(/\s+/g, " ").trim() || activeHref || "";
}

export function toComparableTime(value: string | undefined): number | null {
  const numericTime = Date.parse(value || "");
  return Number.isFinite(numericTime) ? numericTime : null;
}

export function toComparableNumericId(value: string | undefined): number | null {
  if (!/^\d+$/.test(String(value || ""))) {
    return null;
  }

  const numericId = Number(value);
  return Number.isFinite(numericId) ? numericId : null;
}

export function clearDomSelection(): void {
  const selection = window.getSelection?.();
  if (selection && selection.rangeCount) {
    selection.removeAllRanges();
  }
}

export function isToolbarNode(node: EventTarget | null): boolean {
  return Boolean((node as Element)?.closest?.(`.${DOCK_CLASS}`));
}

export function isInteractiveNode(node: EventTarget | null): boolean {
  return Boolean(
    (node as Element)?.closest?.(
      'a, button, input, textarea, summary, select, option, [contenteditable="true"], [data-tsm-selection-control="true"]'
    )
  );
}

export function isWithinMessageRunway(node: EventTarget | null): boolean {
  return Boolean((node as Element)?.closest?.(RUNWAY_SELECTOR));
}

export function isMentionElement(element: Element): boolean {
  const ariaLabel = element?.getAttribute?.("aria-label") || "";
  return (
    ariaLabel.startsWith("Mentioned ") ||
    element?.getAttribute?.("itemtype") === "http://schema.skype.com/Mention"
  );
}
