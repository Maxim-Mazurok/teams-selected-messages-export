import type { LinkContext, MessageSnapshot } from "./types.js";
import { resolveConversationId } from "./api-client.js";
import { getConversationTitle } from "./conversation.js";
import { getCapturedTenantId, getCapturedGroupId } from "./worker-store.js";

const TEAMS_BASE_URL = "https://teams.microsoft.com";
const CHANNEL_PATTERN = /@thread\.(tacv2|skype)$/i;
const CHAT_CONTEXT = encodeURIComponent(JSON.stringify({ contextType: "chat" }));

/**
 * Strip DOM-specific prefixes from message IDs.
 * DOM extraction produces IDs like "content-1773975961503" or "timestamp-1773975961503",
 * but Teams deep links use the bare numeric ID.
 */
function normalizeMessageIdForLink(messageId: string): string {
  return messageId.replace(/^(content-|timestamp-)/, "");
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
    groupId: isChannel ? getCapturedGroupId() : null,
    teamName: isChannel ? extractTeamName() : null,
    channelName: isChannel ? title : null
  };
}

export function buildConversationLink(context: LinkContext): string {
  if (context.isChannel) {
    const encodedConversationId = encodeURIComponent(context.conversationId);
    const encodedChannelName = encodeURIComponent(context.channelName || "Channel");
    const parameters = new URLSearchParams();
    if (context.groupId) {
      parameters.set("groupId", context.groupId);
    }
    if (context.tenantId) {
      parameters.set("tenantId", context.tenantId);
    }
    const query = parameters.toString();
    return `${TEAMS_BASE_URL}/l/channel/${encodedConversationId}/${encodedChannelName}${query ? `?${query}` : ""}`;
  }

  const encodedConversationId = encodeURIComponent(context.conversationId);
  return `${TEAMS_BASE_URL}/l/chat/${encodedConversationId}/conversations?context=${CHAT_CONTEXT}`;
}

export function buildMessageLink(
  context: LinkContext,
  message: MessageSnapshot
): string {
  const normalizedMessageId = normalizeMessageIdForLink(message.id);
  const encodedConversationId = encodeURIComponent(context.conversationId);
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
    return `${TEAMS_BASE_URL}/l/message/${encodedConversationId}/${encodedMessageId}?${parameters.toString()}`;
  }

  return `${TEAMS_BASE_URL}/l/message/${encodedConversationId}/${encodedMessageId}?context=${CHAT_CONTEXT}`;
}

/**
 * Attempt to extract the team name from the DOM sidebar.
 * Channels are nested under their team in the sidebar tree.
 */
function extractTeamName(): string | null {
  // Look for the active channel's parent team in sidebar tree
  const activeChannel = document.querySelector(
    '[role="treeitem"][tabindex="0"][data-testid*="channel-list-item-19:"]'
  );
  if (activeChannel) {
    // Walk up to find the team-level treeitem
    let current: Element | null = activeChannel.parentElement;
    for (let depth = 0; depth < 10 && current; depth++) {
      if (
        current.getAttribute("role") === "treeitem" &&
        !current.getAttribute("data-testid")?.includes("channel-list-item")
      ) {
        const teamNameElement = current.querySelector('[data-tid="team-name"]');
        if (teamNameElement?.textContent?.trim()) {
          return teamNameElement.textContent.trim();
        }
      }
      current = current.parentElement;
    }
  }

  return null;
}
