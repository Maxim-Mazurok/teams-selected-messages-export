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

export function isChannelThread(element: HTMLElement): boolean {
  return element.getAttribute("data-tid") === "channel-pane-message";
}

function buildChannelCompositeContent(threadElement: HTMLElement): HTMLElement {
  const composite = document.createElement("div");

  const responseSurface = threadElement.querySelector('[data-tid="response-surface"]');
  const allBodies = Array.from(threadElement.querySelectorAll('[data-tid="message-body"]'));
  const replyBodies = responseSurface
    ? Array.from(responseSurface.querySelectorAll('[data-tid="message-body"]'))
    : [];
  const replyBodiesSet = new Set(replyBodies);

  const postBody = allBodies.find((body) => !replyBodiesSet.has(body));
  if (postBody) {
    composite.appendChild(postBody.cloneNode(true));
  }

  if (responseSurface) {
    const replyHeaders = Array.from(
      responseSurface.querySelectorAll('[data-tid="reply-message-header"]')
    );

    for (let i = 0; i < replyBodies.length; i++) {
      composite.appendChild(document.createElement("hr"));

      const header = replyHeaders[i];
      if (header) {
        const authorSpan = header.querySelector('span[id^="author-"]');
        const timeElement = header.querySelector('[data-tid="timestamp"], time');
        const author = normalizeText(authorSpan?.textContent || "");
        const timeLabel = normalizeText(timeElement?.textContent || "");

        const attribution = document.createElement("p");
        const strong = document.createElement("strong");
        strong.textContent = [author, timeLabel].filter(Boolean).join(" | ");
        attribution.appendChild(strong);
        composite.appendChild(attribution);
      }

      composite.appendChild(replyBodies[i].cloneNode(true));
    }
  }

  return composite;
}

export function getContentSource(element: HTMLElement, strategy: Strategy | null): HTMLElement {
  if (isChannelThread(element)) {
    return buildChannelCompositeContent(element);
  }

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

export function parseChannelTimeAriaLabel(ariaLabel: string): string {
  const cleaned = ariaLabel.trim();
  if (!cleaned) {
    return "";
  }

  const parsed = Date.parse(cleaned);
  if (Number.isFinite(parsed)) {
    return new Date(parsed).toISOString();
  }

  return "";
}

export function extractSubject(element: HTMLElement): string {
  const subjectElement = element.querySelector('[data-tid="subject-line"]');
  return normalizeText(subjectElement?.textContent || "");
}

export function getTimeMeta(element: HTMLElement, strategy: Strategy | null): TimeMeta {
  const timeElement = strategy?.timeSelector
    ? element.querySelector(strategy.timeSelector)
    : element.querySelector("time");

  if (!timeElement) {
    return { label: "", dateTime: "" };
  }

  const label = normalizeText(timeElement.textContent || "");
  const explicitDatetime = timeElement.getAttribute("datetime") || (timeElement as HTMLTimeElement).dateTime || "";

  if (explicitDatetime) {
    return { label, dateTime: explicitDatetime };
  }

  const ariaLabel = timeElement.getAttribute("aria-label") || "";
  const parsedDatetime = parseChannelTimeAriaLabel(ariaLabel);

  return {
    label: label || normalizeText(ariaLabel),
    dateTime: parsedDatetime
  };
}
