import type { LinkContext, MessageSnapshot } from "./types.js";
import { resolveConversationId } from "./api-client.js";
import { getConversationTitle } from "./conversation.js";
import { getCapturedTenantId, getCapturedGroupId } from "./worker-store.js";

const TEAMS_BASE_URL = "https://teams.cloud.microsoft";
const CHANNEL_PATTERN = /@thread\.(tacv2|skype)$/i;
const CHAT_CONTEXT = encodeURIComponent(JSON.stringify({ contextType: "chat" }));

function normalizeMessageIdForLink(messageId: string): string {
  return messageId.replace(/^(content-|timestamp-|reply-chain-summary-|message-body-)/, "");
}

/**
 * Encode a conversation ID for use in a URL path segment.
 * Unlike encodeURIComponent, this preserves `:` and `@` which Teams expects unencoded.
 */
function encodeConversationIdForPath(conversationId: string): string {
  return encodeURIComponent(conversationId).replace(/%3A/gi, ":").replace(/%40/gi, "@");
}

/**
 * Serialize URLSearchParams using `%20` for spaces instead of `+`.
 * Teams deep links expect percent-encoded spaces.
 */
function formatSearchParameters(parameters: URLSearchParams): string {
  return parameters.toString().replace(/\+/g, "%20");
}

export function isChannelConversationId(conversationId: string): boolean {
  return CHANNEL_PATTERN.test(conversationId);
}

export function buildLinkContext(): LinkContext | null {
  const conversationId = resolveConversationId();
  if (!conversationId) {
    return null;
  }

  const isChannel = isChannelConversationId(conversationId);
  const title = getConversationTitle();

  return {
    conversationId,
    isChannel,
    tenantId: getCapturedTenantId(),
    groupId: isChannel ? (getCapturedGroupId() ?? extractGroupIdFromDom(conversationId)) : null,
    teamName: isChannel ? extractTeamName() : null,
    channelName: isChannel ? title : null
  };
}

export function buildConversationLink(context: LinkContext): string {
  if (context.isChannel) {
    const conversationIdPath = encodeConversationIdForPath(context.conversationId);
    const encodedChannelName = encodeURIComponent(context.channelName || "Channel");
    const parameters = new URLSearchParams();
    if (context.groupId) {
      parameters.set("groupId", context.groupId);
    }
    if (context.tenantId) {
      parameters.set("tenantId", context.tenantId);
    }
    const query = formatSearchParameters(parameters);
    return `${TEAMS_BASE_URL}/l/channel/${conversationIdPath}/${encodedChannelName}${query ? `?${query}` : ""}`;
  }

  const conversationIdPath = encodeConversationIdForPath(context.conversationId);
  return `${TEAMS_BASE_URL}/l/chat/${conversationIdPath}/conversations?context=${CHAT_CONTEXT}`;
}

export function buildMessageLink(
  context: LinkContext,
  message: MessageSnapshot
): string {
  const normalizedMessageId = normalizeMessageIdForLink(message.id);
  const conversationIdPath = encodeConversationIdForPath(context.conversationId);
  const encodedMessageId = encodeURIComponent(normalizedMessageId);

  if (context.isChannel) {
    const parentMessageId = message.threadId
      ? normalizeMessageIdForLink(message.threadId)
      : normalizedMessageId;
    const parameters = new URLSearchParams();
    if (context.tenantId) {
      parameters.set("tenantId", context.tenantId);
    }
    if (context.groupId) {
      parameters.set("groupId", context.groupId);
    }
    parameters.set("parentMessageId", parentMessageId);
    if (context.teamName) {
      parameters.set("teamName", context.teamName);
    }
    if (context.channelName) {
      parameters.set("channelName", context.channelName);
    }
    parameters.set("createdTime", normalizedMessageId);
    return `${TEAMS_BASE_URL}/l/message/${conversationIdPath}/${encodedMessageId}?${formatSearchParameters(parameters)}`;
  }

  return `${TEAMS_BASE_URL}/l/message/${conversationIdPath}/${encodedMessageId}?context=${CHAT_CONTEXT}`;
}

/**
 * Attempt to extract the team name from the DOM sidebar.
 * In the new Teams client, channels are level-3 treeitems inside a [role="group"]
 * whose parent is a level-2 treeitem with data-testid="list-item-teams-and-channels".
 * The team name is the first text child of that parent treeitem.
 */
function extractTeamName(): string | null {
  const activeChannel = document.querySelector(
    '[role="treeitem"][tabindex="0"][data-testid*="channel-list-item-19:"]'
  );
  if (activeChannel) {
    // Walk up through [role="group"] to the parent team treeitem
    const group = activeChannel.closest('[role="group"]');
    const teamTreeitem = group?.parentElement;
    if (
      teamTreeitem?.getAttribute("role") === "treeitem" &&
      teamTreeitem.getAttribute("data-testid") === "list-item-teams-and-channels"
    ) {
      const firstChild = teamTreeitem.children[0];
      const teamName = firstChild?.textContent?.trim();
      if (teamName) {
        return teamName;
      }
    }
  }

  return null;
}

/**
 * Extract the M365 Group ID (Team ID) from DOM elements such as Praise or
 * other adaptive card links that embed groupId in their href.
 */
function extractGroupIdFromDom(conversationId: string): string | null {
  // Links in channel messages (e.g., Praise cards) contain groupId as a URL parameter
  const channelIdEncoded = encodeURIComponent(conversationId);
  const links = document.querySelectorAll<HTMLAnchorElement>(`a[href*="groupId"]`);
  for (const link of links) {
    const href = link.getAttribute("href") || "";
    // Only trust links that reference the current channel
    if (!href.includes(conversationId) && !href.includes(channelIdEncoded)) {
      continue;
    }
    const match = href.match(/groupId[=:]([a-f0-9-]{36})/i) ||
      href.match(/groupId%3D([a-f0-9-]{36})/i) ||
      href.match(/groupId%253D([a-f0-9-]{36})/i);
    if (match) {
      return match[1];
    }
  }

  // Broader fallback: any link with groupId (even without channel reference)
  for (const link of links) {
    const href = link.getAttribute("href") || "";
    const match = href.match(/groupId[=:]([a-f0-9-]{36})/i) ||
      href.match(/groupId%3D([a-f0-9-]{36})/i) ||
      href.match(/groupId%253D([a-f0-9-]{36})/i);
    if (match) {
      return match[1];
    }
  }

  return null;
}
