import type { QuotedReply, ReactionInfo, MessageSnapshot, ExportOptions, LinkContext } from "./types.js";
import { escapeHtml } from "./utilities.js";
import { formatQuotedReplyLabel } from "./markdown-renderer.js";
import { buildConversationLink, buildMessageLink } from "./link-builder.js";

function renderQuotedReplyHtml(quote: QuotedReply | null): string {
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

function renderReactionsHtml(reactions: ReactionInfo[] | undefined): string {
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

export function renderHtmlDocument(
  messages: MessageSnapshot[],
  meta: { title: string; sourceUrl: string; exportedAt: string; scope?: string; options?: ExportOptions; linkContext?: LinkContext | null }
): string {
  const scope = meta.scope || "selection";
  const includeLinks = meta.options?.includeLinks === true && meta.linkContext;
  const summaryMessage =
    scope === "full-chat"
      ? `${messages.length} message${messages.length === 1 ? "" : "s"} captured from the full chat history.`
      : `${messages.length} message${messages.length === 1 ? "" : "s"} selected.`;

  const conversationLinkHtml = includeLinks && meta.linkContext
    ? `\n        <p><a href="${escapeHtml(buildConversationLink(meta.linkContext))}" class="teams-link">Open in Teams ↗</a></p>`
    : "";

  const articles = messages
    .map((message) => {
      const safeAuthor = escapeHtml(message.author);
      const safeTime = escapeHtml(message.timeLabel || message.dateTime || "");
      const quoteHtml = renderQuotedReplyHtml(message.quote);
      const reactionsHtml = renderReactionsHtml(message.reactions);
      const subjectHtml = message.subject ? `<h3>${escapeHtml(message.subject)}</h3>` : "";
      const replyClass = message.isReply ? " thread-reply" : "";
      const messageLinkHtml = includeLinks && meta.linkContext
        ? ` <a href="${escapeHtml(buildMessageLink(meta.linkContext, message))}" class="message-link" title="Open in Teams">↗</a>`
        : "";
      return `
          <article class="message${replyClass}">
            <header>
              <strong>${message.isReply ? "↳ " : ""}${safeAuthor}${messageLinkHtml}</strong>
              <time datetime="${escapeHtml(message.dateTime || "")}">${safeTime}</time>
            </header>
            <section class="body">${subjectHtml}${quoteHtml}${message.html || `<p>${escapeHtml(message.plainText)}</p>`}${reactionsHtml}</section>
          </article>
        `;
    })
    .join("\n");

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(meta.title)} - Teams export</title>
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
      .thread-reply {
        margin-left: 32px;
        border-left: 3px solid #0d61d8;
        background: #f9fbfe;
      }
      .teams-link {
        display: inline-block;
        padding: 6px 14px;
        border-radius: 999px;
        background: #0d61d8;
        color: #fff;
        text-decoration: none;
        font-size: 13px;
        font-weight: 600;
      }
      .teams-link:hover {
        background: #0a4fa8;
      }
      .message-link {
        color: #0d61d8;
        text-decoration: none;
        font-size: 12px;
        margin-left: 6px;
        opacity: 0.6;
      }
      .message-link:hover {
        opacity: 1;
        text-decoration: underline;
      }
    </style>
  </head>
  <body>
    <main>
      <section class="summary">
        <h1>${escapeHtml(meta.title)}</h1>
        <p>Exported from Microsoft Teams on ${escapeHtml(meta.exportedAt)}.</p>
        <p>${escapeHtml(summaryMessage)}</p>
        <p>Source: <a href="${escapeHtml(meta.sourceUrl)}">${escapeHtml(meta.sourceUrl)}</a></p>${conversationLinkHtml}
      </section>
      ${articles}
    </main>
  </body>
</html>`;
}
