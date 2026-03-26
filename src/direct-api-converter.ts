/**
 * Converts Teams direct API response messages into MessageSnapshot objects
 * compatible with the existing markdown and HTML renderers.
 */

import type { MessageSnapshot, QuotedReply, ReactionInfo } from "./types.js";
import type { TeamsApiMessage, TeamsApiMember, TeamsApiEmotion } from "./direct-api-client.js";

/** Map from Teams emotion keys to display emoji + name */
const EMOTION_MAP: Record<string, { emoji: string; name: string }> = {
  like: { emoji: "👍", name: "like" },
  heart: { emoji: "❤️", name: "heart" },
  laugh: { emoji: "😂", name: "laugh" },
  surprised: { emoji: "😮", name: "surprised" },
  sad: { emoji: "😢", name: "sad" },
  angry: { emoji: "😡", name: "angry" },
  fistbump: { emoji: "🤜", name: "fist bump" },
  heartlightblue: { emoji: "💙", name: "blue heart" },
  yes: { emoji: "✅", name: "yes" },
  no: { emoji: "❌", name: "no" },
  clap: { emoji: "👏", name: "clap" },
  fire: { emoji: "🔥", name: "fire" },
  mindblown: { emoji: "🤯", name: "mind blown" },
  party: { emoji: "🎉", name: "party" },
  eyes: { emoji: "👀", name: "eyes" },
  pensive: { emoji: "😔", name: "pensive" },
  rocket: { emoji: "🚀", name: "rocket" },
  hundred: { emoji: "💯", name: "hundred" },
  plusone: { emoji: "👍", name: "+1" },
};

/**
 * Convert a batch of API messages to MessageSnapshot objects.
 *
 * @param apiMessages - Messages from the Teams REST API (should be sorted newest-first from API)
 * @param members - Optional member list for resolving user IDs to display names
 * @returns MessageSnapshot[] sorted oldest-first (chronological order for rendering)
 */
export function convertApiMessagesToSnapshots(
  apiMessages: TeamsApiMessage[],
  members: TeamsApiMember[] = [],
): MessageSnapshot[] {
  // Build a member lookup by MRI (e.g., "8:orgid:uuid")
  const memberLookup = new Map<string, string>();
  for (const member of members) {
    memberLookup.set(member.id, member.displayName);
  }

  // Also build a lookup from message authors
  for (const message of apiMessages) {
    if (message.from && message.displayName) {
      // The 'from' field is a URL like "https://...;messageid=xxx"
      // The actual MRI is in the URL path. Extract it.
      const mriMatch = message.from.match(/(8:orgid:[a-f0-9-]+)/);
      if (mriMatch) {
        memberLookup.set(mriMatch[1], message.displayName);
      }
    }
  }

  // Build a message lookup for quoted replies
  const messageLookup = new Map<string, TeamsApiMessage>();
  for (const message of apiMessages) {
    messageLookup.set(message.id, message);
  }

  // Filter to text messages only (skip system events, topic updates, etc.)
  const textMessages = apiMessages.filter(
    (message) =>
      (message.messageType === "RichText/Html" || message.messageType === "Text") &&
      !message.isDeleted,
  );

  // Sort chronologically (oldest first) for rendering
  const sorted = [...textMessages].sort(
    (first, second) =>
      new Date(first.originalArrivalTime).getTime() -
      new Date(second.originalArrivalTime).getTime(),
  );

  return sorted.map((message, index) =>
    convertSingleMessage(message, index, memberLookup, messageLookup),
  );
}

function convertSingleMessage(
  message: TeamsApiMessage,
  index: number,
  memberLookup: Map<string, string>,
  messageLookup: Map<string, TeamsApiMessage>,
): MessageSnapshot {
  const dateTime = message.originalArrivalTime || message.composeTime;
  const timeLabel = formatTimeLabel(dateTime);

  // Parse HTML content for plain text
  const plainText = stripHtml(message.content);

  // Build quoted reply if present
  const quote = buildQuotedReply(message, messageLookup);

  // Build reactions
  const reactions = buildReactions(message.emotions, memberLookup);

  // Convert HTML to basic markdown
  const markdown = htmlToBasicMarkdown(message.content);

  return {
    id: message.id,
    index,
    author: message.displayName || "Unknown",
    timeLabel,
    dateTime,
    subject: message.subject ?? "",
    quote,
    reactions,
    html: message.content,
    markdown,
    plainText,
  };
}

function formatTimeLabel(isoString: string): string {
  try {
    const date = new Date(isoString);
    const day = date.toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
    const time = date.toLocaleTimeString("en-GB", {
      hour: "2-digit",
      minute: "2-digit",
    });
    return `${day}, ${time}`;
  } catch {
    return isoString;
  }
}

function stripHtml(html: string): string {
  return html
    .replace(/<blockquote[\s\S]*?<\/blockquote>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/\s+/g, " ")
    .trim();
}

function htmlToBasicMarkdown(html: string): string {
  return html
    .replace(/<blockquote[\s\S]*?<\/blockquote>/gi, (match) => {
      const text = stripHtml(match);
      return text
        .split("\n")
        .map((line) => `> ${line}`)
        .join("\n");
    })
    .replace(/<strong>([\s\S]*?)<\/strong>/gi, "**$1**")
    .replace(/<em>([\s\S]*?)<\/em>/gi, "*$1*")
    .replace(/<code>([\s\S]*?)<\/code>/gi, "`$1`")
    .replace(/<a[^>]+href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, "[$2]($1)")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>\s*<p>/gi, "\n\n")
    .replace(/<\/?(p|div|span)[^>]*>/gi, "")
    .replace(/<img[^>]+alt="([^"]*)"[^>]*>/gi, "[$1]")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .trim();
}

function buildQuotedReply(
  message: TeamsApiMessage,
  messageLookup: Map<string, TeamsApiMessage>,
): QuotedReply | null {
  if (!message.quotedMessageId) return null;

  // Extract quote info from HTML content
  const content = message.content;
  const authorMatch = content.match(
    /<strong[^>]*itemprop="mri"[^>]*>([\s\S]*?)<\/strong>/i,
  );
  const previewMatch = content.match(
    /<p[^>]*itemprop="preview"[^>]*>([\s\S]*?)<\/p>/i,
  );
  const timeMatch = content.match(
    /<span[^>]*itemprop="time"[^>]*itemid="(\d+)"[^>]*>/i,
  );

  const author = authorMatch ? stripHtml(authorMatch[1]) : "Unknown";
  const text = previewMatch ? stripHtml(previewMatch[1]) : "";

  let timeLabel = "";
  if (timeMatch) {
    // The itemid is a timestamp
    const timestamp = Number(timeMatch[1]);
    if (timestamp > 0) {
      timeLabel = formatTimeLabel(new Date(timestamp).toISOString());
    }
  } else {
    // Try to get from the quoted message
    const quotedMessage = messageLookup.get(message.quotedMessageId);
    if (quotedMessage) {
      timeLabel = formatTimeLabel(quotedMessage.originalArrivalTime);
    }
  }

  return { author, timeLabel, text };
}

function buildReactions(
  emotions: TeamsApiEmotion[],
  memberLookup: Map<string, string>,
): ReactionInfo[] {
  if (!emotions || emotions.length === 0) return [];

  return emotions.map((emotion) => {
    const mapped = EMOTION_MAP[emotion.key] ?? {
      emoji: "👍",
      name: emotion.key,
    };

    const actors = emotion.users.map((user) => {
      const displayName = memberLookup.get(user.mri);
      if (displayName) return displayName;

      // Try to extract a readable ID from the MRI
      const match = user.mri.match(/8:orgid:(.+)/);
      return match ? match[1].slice(0, 8) : user.mri;
    });

    return {
      emoji: mapped.emoji,
      name: mapped.name,
      count: emotion.users.length,
      actors,
    };
  });
}
