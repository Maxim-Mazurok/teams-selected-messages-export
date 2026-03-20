/**
 * Worker Store
 *
 * Listens for bridge messages posted by the document_start MAIN-world hook
 * (extension-src/worker-hook.js) and stores the resulting structured data so
 * it can be used to enrich DOM-extracted message records.
 *
 * Communication channel:
 *   worker-hook.js  →  window.postMessage({ source: "teams-export-worker-hook", ... })
 *   worker-store.ts →  window.addEventListener("message", ...)
 *
 * Data stored:
 *   workerMessageMap  Map<messageId, WorkerCapturedMessage>
 *   memberDisplayNameMap  Map<userId, displayName>
 */

import type { WorkerCapturedMessage } from "./types.js";
import { log } from "./state.js";

const BRIDGE_SOURCE = "teams-export-worker-hook";

/** All messages received from worker responses, keyed by message ID. */
const workerMessageMap = new Map<string, WorkerCapturedMessage>();

/** User-ID to display-name map, built from fromUser fields and member queries. */
const memberDisplayNameMap = new Map<string, string>();

let totalBatchesReceived = 0;
let totalMessagesReceived = 0;

/** Auth token captured from fetch/XHR interception in the MAIN world. */
let capturedSkypeToken: string | null = null;
let capturedRegion: string | null = null;
let capturedTokenType: "skypetoken" | "bearer" = "skypetoken";

/** Tenant ID extracted from the JWT access token. */
let capturedTenantId: string | null = null;

/** M365 Group IDs (Team IDs) extracted from localStorage. */
const capturedGroupIds: string[] = [];

/** Conversation ID extracted from worker message batches. */
let lastSeenConversationId: string | null = null;

// ── Public accessors ──────────────────────────────────────────────────────────

/**
 * Normalise DOM-extracted message IDs to the bare numeric/string IDs used by the worker.
 *
 * Teams DOM child element IDs follow patterns such as:
 *   - "content-1773218944080"     → "1773218944080"
 *   - "timestamp-1773218944080"   → "1773218944080"
 *
 * Worker IDs are always the bare value ("1773218944080"), so we strip known
 * prefixes when looking up the map.
 */
function normalizeMessageId(rawId: string): string {
  return rawId.replace(/^(content-|timestamp-)/, "");
}

export function getWorkerMessage(messageId: string): WorkerCapturedMessage | undefined {
  return workerMessageMap.get(messageId) ?? workerMessageMap.get(normalizeMessageId(messageId));
}

export function resolveUserId(userId: string): string | undefined {
  return memberDisplayNameMap.get(userId);
}

export function getWorkerStats(): { batches: number; messages: number; members: number } {
  return {
    batches: totalBatchesReceived,
    messages: totalMessagesReceived,
    members: memberDisplayNameMap.size
  };
}

export function isWorkerHookInstalled(): boolean {
  return Boolean((window as Window & { __teamsExportWorkerHookInstalled?: boolean }).__teamsExportWorkerHookInstalled);
}

export function getCapturedToken(): string | null {
  return capturedSkypeToken;
}

export function getCapturedRegion(): string | null {
  return capturedRegion;
}

export function getCapturedTokenType(): "skypetoken" | "bearer" {
  return capturedTokenType;
}

export function getLastSeenConversationId(): string | null {
  return lastSeenConversationId;
}

export function getCapturedTenantId(): string | null {
  return capturedTenantId;
}

/**
 * Get the best-guess Group ID (Team ID) for the current channel.
 * Returns the first captured group ID, or null if none available.
 */
export function getCapturedGroupId(): string | null {
  return capturedGroupIds.length > 0 ? capturedGroupIds[0] : null;
}

// ── Internal helpers ──────────────────────────────────────────────────────────

function registerMember(userId: string, displayName: string | null): void {
  if (!userId || memberDisplayNameMap.has(userId)) return;
  if (displayName) {
    memberDisplayNameMap.set(userId, displayName);
  }
}

function handleMessageBatch(payload: {
  operationName: string | null;
  requestId: string | null;
  dataSource: string | null;
  workerUrl: string;
  messages: WorkerCapturedMessage[];
}): void {
  if (!Array.isArray(payload.messages)) return;

  let newCount = 0;
  for (const message of payload.messages) {
    if (!message.id) continue;

    const isNew = !workerMessageMap.has(message.id);
    workerMessageMap.set(message.id, message);
    if (isNew) newCount++;

    // Register author as a known member if we haven't seen them before
    if (message.authorId && message.author) {
      const isNewMember = !memberDisplayNameMap.has(message.authorId);
      registerMember(message.authorId, message.author);
      if (isNewMember && memberDisplayNameMap.has(message.authorId)) {
        document.documentElement.dataset.teamsExportWorkerMembers = String(memberDisplayNameMap.size);
      }
    }

    // Track conversation ID from message records
    if (message.conversationId && !lastSeenConversationId) {
      lastSeenConversationId = message.conversationId;
      document.documentElement.dataset.teamsExportConversationId = lastSeenConversationId;
    }
  }

  totalBatchesReceived++;
  totalMessagesReceived += newCount;

  if (newCount > 0) {
    log(
      `[worker-store] +${newCount} messages from ${payload.operationName || "unknown"} ` +
        `(source: ${payload.dataSource || "?"}, total: ${workerMessageMap.size})`
    );
    // Update diagnostic DOM attribute so the worker POC validation can verify
    // bridge messages are being received without needing main-world window access.
    document.documentElement.dataset.teamsExportWorkerMessages = String(workerMessageMap.size);
  }
}

function handleMeIdentity(payload: { id: string; displayName: string | null }): void {
  if (!payload.id) return;
  registerMember(payload.id, payload.displayName);
  log(`[worker-store] me = ${payload.displayName || payload.id}`);
}

function handleMembers(payload: { members: Array<{ id: string; displayName: string | null }> }): void {
  if (!Array.isArray(payload.members)) return;
  let newCount = 0;
  for (const member of payload.members) {
    if (!member.id) continue;
    const isNew = !memberDisplayNameMap.has(member.id);
    registerMember(member.id, member.displayName);
    if (isNew && member.displayName) newCount++;
  }
  if (newCount > 0) {
    log(`[worker-store] +${newCount} members (total: ${memberDisplayNameMap.size})`);
    document.documentElement.dataset.teamsExportWorkerMembers = String(memberDisplayNameMap.size);
  }
}

function handleAuthToken(payload: { token: string; region: string | null; tokenType?: string; source: string; tenantId?: string | null }): void {
  if (!payload.token || payload.token.length < 100) return;

  const isNew = capturedSkypeToken !== payload.token;
  capturedSkypeToken = payload.token;
  if (payload.region) {
    capturedRegion = payload.region;
  }
  if (payload.tokenType === "bearer") {
    capturedTokenType = "bearer";
  }
  if (payload.tenantId) {
    capturedTenantId = payload.tenantId;
  }

  if (isNew) {
    log(
      `[worker-store] auth token captured (type: ${capturedTokenType}, region: ${capturedRegion ?? "unknown"}, ` +
        `tenant: ${capturedTenantId ?? "unknown"}, source: ${payload.source}, length: ${payload.token.length})`
    );
  }
}

function handleGroupIds(payload: { groupIds: string[] }): void {
  if (!Array.isArray(payload.groupIds)) return;
  for (const groupId of payload.groupIds) {
    if (groupId && !capturedGroupIds.includes(groupId)) {
      capturedGroupIds.push(groupId);
    }
  }
  if (capturedGroupIds.length > 0) {
    log(`[worker-store] captured ${capturedGroupIds.length} group ID(s)`);
  }
}

// ── Event listener setup ──────────────────────────────────────────────────────

function onWindowMessage(event: MessageEvent): void {
  // Only accept messages originating from this same page window, not from iframes
  if (event.source !== window) return;

  const data = event.data;
  if (!data || typeof data !== "object") return;
  if (data.source !== BRIDGE_SOURCE) return;

  const { type, payload } = data as { type: string; payload: unknown };
  if (!type || !payload || typeof payload !== "object") return;

  try {
    if (type === "message-batch") {
      handleMessageBatch(payload as Parameters<typeof handleMessageBatch>[0]);
    } else if (type === "me") {
      handleMeIdentity(payload as Parameters<typeof handleMeIdentity>[0]);
    } else if (type === "members") {
      handleMembers(payload as Parameters<typeof handleMembers>[0]);
    } else if (type === "auth-token") {
      handleAuthToken(payload as { token: string; region: string | null; tokenType?: string; source: string; tenantId?: string | null });
    } else if (type === "group-ids") {
      handleGroupIds(payload as { groupIds: string[] });
    }
  } catch (error) {
    log(`[worker-store] error processing bridge event "${type}":`, error);
  }
}

/** Call once from main.ts to start listening for bridge events. */
export function initWorkerStore(): void {
  window.addEventListener("message", onWindowMessage);
  log(
    `[worker-store] listening for worker bridge events`,
    isWorkerHookInstalled()
      ? "(hook confirmed installed)"
      : "(hook NOT detected — was document_start injection missed?)"
  );
}
