import type { Strategy, QuotedReply, ReactionInfo, TimeMeta } from "./types.js";
import { CONTENT_PRUNE_SELECTOR } from "./constants.js";
import { normalizeText, isMentionElement } from "./utilities.js";

export function formatMentionLabel(element: Element): string {
  const ariaLabel = element.getAttribute("aria-label") || "";
  const explicitName = ariaLabel.startsWith("Mentioned ") ? ariaLabel.replace(/^Mentioned\s+/, "") : "";
  const text = normalizeText(element.textContent || "");
  const name = explicitName || text;
  return name ? `@${name}` : text;
}

export function isEmojiImage(image: Element): boolean {
  return Boolean(
    image.closest('[data-tid="emoticon-renderer"]') ||
      image.getAttribute("itemtype") === "http://schema.skype.com/Emoji"
  );
}

export function isDecorativeImage(image: Element): boolean {
  return Boolean(
    image.closest('[data-tid*="reaction"]') ||
      image.closest('[data-tid="message-avatar"]') ||
      image.closest('[data-tid="quoted-reply-card"]')
  );
}

export function createImagePlaceholder(label: string): HTMLSpanElement {
  const placeholder = document.createElement("span");
  placeholder.textContent = label;
  placeholder.dataset.tsmPlaceholder = "image";
  return placeholder;
}

export function pruneClone(root: HTMLElement): HTMLElement {
  root.querySelectorAll(CONTENT_PRUNE_SELECTOR).forEach((node) => node.remove());
  return root;
}

export function getContentSource(element: HTMLElement, strategy: Strategy | null): HTMLElement {
  return ((strategy?.contentSelector && element.querySelector(strategy.contentSelector)) as HTMLElement) || element;
}

export function normalizeContentClone(root: HTMLElement): HTMLElement {
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

export function getPreparedContentClone(element: HTMLElement, strategy: Strategy | null): HTMLElement {
  const clone = pruneClone(getContentSource(element, strategy).cloneNode(true) as HTMLElement);
  return normalizeContentClone(clone);
}

export function extractBodyHtml(element: HTMLElement, strategy: Strategy | null): string {
  const clone = getPreparedContentClone(element, strategy);
  return clone.innerHTML.trim();
}

export function extractPlainText(element: HTMLElement, strategy: Strategy | null): string {
  return normalizeText(getPreparedContentClone(element, strategy).textContent || "");
}

export function extractQuotedReply(element: HTMLElement): QuotedReply | null {
  const card = element.querySelector('[data-tid="quoted-reply-card"]');
  if (!card) {
    return null;
  }

  const timeLabel = normalizeText(
    card.querySelector('[data-tid="quoted-reply-timestamp"]')?.textContent || ""
  );
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

export function extractReactionActors(button: Element, labelText: string): string[] {
  const candidates = [
    button.getAttribute("aria-label"),
    button.getAttribute("title"),
    labelText
  ]
    .map((value) => normalizeText(value || ""))
    .filter(Boolean);

  const actors: string[] = [];

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

export function extractReactions(element: HTMLElement): ReactionInfo[] {
  return Array.from(element.querySelectorAll('[data-tid="diverse-reaction-pill-button"]'))
    .map((button) => {
      const labelId = button.getAttribute("aria-labelledby");
      const labelText = normalizeText(
        (labelId && document.getElementById(labelId)?.textContent) || button.textContent || ""
      );
      const match = labelText.match(/(\d+)\s+(.+?)\s+reactions?\.?/i);
      const emoji = button.querySelector("img[alt]")?.getAttribute("alt") || "";
      const count = Number(match?.[1] || 1);
      const name = normalizeText(match?.[2] || labelText.replace(/\.$/, ""));
      const actors = extractReactionActors(button, labelText);

      if (!emoji && !name) {
        return null;
      }

      return { emoji, name, count, actors };
    })
    .filter((reaction): reaction is ReactionInfo => reaction !== null);
}

export function getTimeMeta(element: HTMLElement, strategy: Strategy | null): TimeMeta {
  const timeElement = strategy?.timeSelector
    ? element.querySelector(strategy.timeSelector)
    : element.querySelector("time");

  if (!timeElement) {
    return { label: "", dateTime: "" };
  }

  return {
    label: normalizeText(timeElement.textContent || ""),
    dateTime: timeElement.getAttribute("datetime") || (timeElement as HTMLTimeElement).dateTime || ""
  };
}
