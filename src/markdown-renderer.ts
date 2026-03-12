import type { Strategy, QuotedReply, ReactionInfo, MessageSnapshot } from "./types.js";
import { normalizeText, inlineMarkdown, isMentionElement } from "./utilities.js";
import {
  formatMentionLabel,
  isEmojiImage,
  isDecorativeImage,
  getPreparedContentClone
} from "./content-extraction.js";

export function formatQuotedReplyLabel(quote: QuotedReply): string {
  const parts: string[] = [];
  if (quote.author) {
    parts.push(quote.author);
  }
  if (quote.timeLabel) {
    parts.push(quote.timeLabel);
  }

  return parts.length ? `Replying to ${parts.join(" | ")}` : "Replying to";
}

export function renderQuotedReplyMarkdown(quote: QuotedReply | null): string {
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

export function renderReactionsMarkdown(reactions: ReactionInfo[] | undefined): string {
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

function nodeToMarkdown(node: Node, context: Record<string, unknown> = {}): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || "";
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return "";
  }

  const element = node as Element;
  const tag = element.tagName.toLowerCase();
  const children = Array.from(element.childNodes)
    .map((child) => nodeToMarkdown(child, context))
    .join("");

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

  if (tag === "hr") {
    return "\n---\n";
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
    const text = inlineMarkdown(children, (element as HTMLAnchorElement).href);
    const href = (element as HTMLAnchorElement).href || "";
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
    if (
      tag === "div" &&
      (isMentionElement(element) || (element as HTMLElement).dataset.tsmMention === "true")
    ) {
      return children.trim();
    }

    return `${children.trim()}\n\n`;
  }

  return children;
}

function mergeConsecutiveMentions(markdown: string): string {
  // Merge consecutive mentions that appear to be first-name + last-name pairs
  // Pattern: @FirstName @LastName -> @FirstName LastName
  // Also handle comma-separated: @FirstName @LastName, -> @FirstName LastName,
  let result = markdown;
  // Apply twice to catch multiple mentions in same text (e.g., "@Leo @Shchurov and @Alan @Markus")
  result = result.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3");
  result = result.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)$/gm, "@$1 $2");
  result = result.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3");
  return result;
}

export function elementToMarkdown(
  element: HTMLElement,
  strategy: Strategy | null,
  fallbackText: string
): string {
  if (!element) {
    return fallbackText;
  }

  const clone = getPreparedContentClone(element, strategy);
  const markdown = Array.from(clone.childNodes)
    .map((node) => nodeToMarkdown(node))
    .join("")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  const mergedMarkdown = mergeConsecutiveMentions(markdown);
  return mergedMarkdown || fallbackText;
}

export function renderMarkdown(
  messages: MessageSnapshot[],
  meta: { title: string; sourceUrl: string; exportedAt: string; scope?: string }
): string {
  const scope = meta.scope || "selection";
  const countLabel =
    scope === "full-chat"
      ? `Full chat history messages: ${messages.length}`
      : `Selected messages: ${messages.length}`;

  const lines = [
    `# ${meta.title}`,
    "",
    `- Exported from Microsoft Teams on ${meta.exportedAt}`,
    `- Source URL: ${meta.sourceUrl}`,
    `- ${countLabel}`,
    ""
  ];

  messages.forEach((message) => {
    const headingBits = [message.author];
    if (message.timeLabel || message.dateTime) {
      headingBits.push(message.timeLabel || message.dateTime);
    }

    lines.push(`## ${headingBits.join(" | ")}`);
    lines.push("");
    if (message.subject) {
      lines.push(`**${message.subject}**`);
      lines.push("");
    }
    if (message.quote?.text) {
      lines.push(renderQuotedReplyMarkdown(message.quote));
      lines.push("");
    }
    lines.push(message.markdown || message.plainText);
    if (message.reactions?.length) {
      lines.push("");
      lines.push(renderReactionsMarkdown(message.reactions));
    }
    lines.push("");
  });

  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
}
