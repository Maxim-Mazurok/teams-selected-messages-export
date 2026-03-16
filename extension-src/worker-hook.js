/**
 * Worker Hook — document_start, MAIN world
 *
 * Patches window.Worker before Microsoft Teams creates its precompiled web worker
 * so that all worker message traffic is intercepted and summarised.
 *
 * Relevant workers created by Teams:
 *   /v2/worker/precompiled-web-worker-*.js  (precore-worker, sends GraphQL)
 *
 * Bridge events are sent to the isolated content script via window.postMessage
 * using the envelope:
 *   { source: "teams-export-worker-hook", type: <string>, payload: <object> }
 *
 * Types bridged:
 *   "message-batch"  — a batch of chat messages from a worker response
 *   "me"             — own user identity from singleMe / me query
 *   "members"        — member list from a chat/members query
 */
(function workerHookIIFE() {
  "use strict";

  const BRIDGE_SOURCE = "teams-export-worker-hook";

  // ── JSON helpers ──────────────────────────────────────────────────────────

  function tryParseJson(value) {
    if (typeof value !== "string") return null;
    try {
      return JSON.parse(value);
    } catch {
      return null;
    }
  }

  // ── Message summarisation ─────────────────────────────────────────────────

  function summarizeEmotionsSummary(summary) {
    if (!Array.isArray(summary)) return null;
    return summary.map(function (entry) {
      return { key: entry.key || null, count: entry.count !== undefined ? entry.count : null };
    });
  }

  function summarizeEmotionUsers(collection) {
    if (!Array.isArray(collection)) return null;
    return collection.map(function (entry) {
      return {
        key: entry.commonKey || entry.key || null,
        userIds: Array.isArray(entry.users)
          ? entry.users.map(function (user) { return user.userId || user.id || null; }).filter(Boolean)
          : []
      };
    });
  }

  function summarizeMessage(message) {
    if (!message || !message.id) return null;
    return {
      id: String(message.id),
      clientMessageId: message.clientMessageId || null,
      originalArrivalTime: message.originalArrivalTime || null,
      author: (message.fromUser && message.fromUser.displayName) || message.imDisplayName || null,
      authorId:
        (message.fromUser && message.fromUser.id) ||
        message.fromUserId ||
        message.from ||
        null,
      messageType: message.messageType || null,
      content: message.content || null,
      subject: message.subject || null,
      threadType: message.threadType || null,
      isDeleted: message.isDeleted || false,
      editedTime: message.updatedTime || null,
      emotionsSummary: summarizeEmotionsSummary(message.emotionsSummary),
      emotions: summarizeEmotionUsers(message.emotions),
      diverseEmotions: summarizeEmotionUsers(
        message.diverseEmotions && message.diverseEmotions.items
          ? message.diverseEmotions.items
          : message.diverseEmotions
      ),
      quotedMessages: message.quotedMessages || null,
    };
  }

  // ── Response data extraction ───────────────────────────────────────────────

  /**
   * Pull a message array out of the various GraphQL response shapes Teams uses.
   * Returns null when the payload does not contain messages.
   */
  function extractMessagesFromResponseData(responseData) {
    if (!responseData || typeof responseData !== "object") return null;

    // ComponentsChatQueriesMessageListQuery / ComponentsChatQueriesMissingMessagesQuery
    if (
      responseData.messages &&
      Array.isArray(responseData.messages.messages)
    ) {
      return responseData.messages.messages;
    }

    // live push via fetchRecentMessagesAfter
    if (Array.isArray(responseData.fetchRecentMessagesAfter)) {
      return responseData.fetchRecentMessagesAfter;
    }

    // DataResolversBrowserChatServiceEventsChatServiceBatchEventSubscription
    if (Array.isArray(responseData.chatServiceEvents)) {
      var events = responseData.chatServiceEvents;
      var messages = [];
      for (var i = 0; i < events.length; i++) {
        var event = events[i];
        if (Array.isArray(event.messages)) {
          for (var j = 0; j < event.messages.length; j++) {
            messages.push(event.messages[j]);
          }
        }
        // single message in a chatServiceEvent subscription push
        if (event.message) {
          messages.push(event.message);
        }
      }
      if (messages.length > 0) return messages;
    }

    // chatServiceEvent subscription (legacy shape)
    if (responseData.chatServiceEvent) {
      var cse = responseData.chatServiceEvent;
      if (Array.isArray(cse.messages)) return cse.messages;
      if (cse.message) return [cse.message];
    }

    return null;
  }

  /**
   * Extract "me" identity from singleMe / me response shapes.
   */
  function extractMeFromResponseData(responseData) {
    if (!responseData) return null;
    var candidate = responseData.singleMe || responseData.me;
    if (!candidate) return null;
    var id = candidate.id || candidate.userId;
    if (!id) return null;
    return {
      id: String(id),
      displayName: candidate.displayName || candidate.name || null
    };
  }

  /**
   * Extract a member list from chatMembers / members / chatParticipants response shapes.
   */
  function extractMembersFromResponseData(responseData) {
    if (!responseData) return null;
    var raw =
      responseData.chatMembers ||
      responseData.members ||
      responseData.chatParticipants ||
      (responseData.chat && (responseData.chat.members || responseData.chat.participants));
    if (!Array.isArray(raw) || raw.length === 0) return null;
    var members = [];
    for (var i = 0; i < raw.length; i++) {
      var member = raw[i];
      var id = member.id || member.userId || member.mri;
      if (!id) continue;
      members.push({
        id: String(id),
        displayName: member.displayName || member.name || null
      });
    }
    return members.length > 0 ? members : null;
  }

  // ── Bridge helpers ─────────────────────────────────────────────────────────

  function postBridge(type, payload) {
    try {
      window.postMessage({ source: BRIDGE_SOURCE, type: type, payload: payload }, "*");
    } catch {
      // Ignore postMessage serialisation errors (circular refs, etc.)
    }
  }

  // ── Worker message processing ──────────────────────────────────────────────

  function processWorkerResponse(rawData, workerUrl) {
    var payload = typeof rawData === "object" && rawData !== null
      ? rawData
      : tryParseJson(rawData);
    if (!payload || typeof payload !== "object") return;

    // Teams GraphQL responses are wrapped: { requestId, response: { data }, extensions }
    var responseData = (payload.response && payload.response.data) || null;
    if (!responseData) return;

    var operationName =
      payload.operationName ||
      (payload.payload && payload.payload.operationName) ||
      null;

    // ── Messages ─────────────────────────────────────────────────────────────
    var rawMessages = extractMessagesFromResponseData(responseData);
    if (rawMessages && rawMessages.length > 0) {
      var summarized = [];
      for (var i = 0; i < rawMessages.length; i++) {
        var summarizedMessage = summarizeMessage(rawMessages[i]);
        if (summarizedMessage) summarized.push(summarizedMessage);
      }
      if (summarized.length > 0) {
        postBridge("message-batch", {
          operationName: operationName,
          requestId: payload.requestId || null,
          dataSource: (payload.extensions && payload.extensions.dataSource) || null,
          workerUrl: String(workerUrl),
          messages: summarized
        });
      }
    }

    // ── Me identity ───────────────────────────────────────────────────────────
    var meData = extractMeFromResponseData(responseData);
    if (meData) {
      postBridge("me", meData);
    }

    // ── Member list ───────────────────────────────────────────────────────────
    var membersData = extractMembersFromResponseData(responseData);
    if (membersData) {
      postBridge("members", { members: membersData });
    }
  }

  // ── Worker constructor patch ───────────────────────────────────────────────

  var OriginalWorker = window.Worker;

  function WorkerWrapper(workerUrl, options) {
    var worker = options
      ? new OriginalWorker(workerUrl, options)
      : new OriginalWorker(workerUrl);

    // Intercept responses coming back from the worker
    worker.addEventListener("message", function (event) {
      try {
        processWorkerResponse(event.data, workerUrl);
      } catch {
        // Never let our hook crash the page
      }
    });

    return worker;
  }

  WorkerWrapper.prototype = OriginalWorker.prototype;
  Object.setPrototypeOf(WorkerWrapper, OriginalWorker);
  window.Worker = WorkerWrapper;

  // Signal that this hook is installed so the content script can verify it
  window.__teamsExportWorkerHookInstalled = true;

  // eslint-disable-next-line no-console
  console.debug("[teams-export] worker hook installed");
})();
