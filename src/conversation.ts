import { normalizeText } from "./utilities.js";

export function getConversationTitle(): string {
  const candidates = [
    '[data-tid="chat-title"]',
    '[data-tid="chat-header-title"]',
    '[data-tid="thread-header-title"]',
    "main h1",
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

export function getConversationKey(): string {
  const title = getConversationTitle();
  const documentTitle = normalizeText(document.title || "");
  const path = `${location.pathname}${location.search}${location.hash}`;
  return [title, documentTitle, path].filter(Boolean).join(" | ");
}
