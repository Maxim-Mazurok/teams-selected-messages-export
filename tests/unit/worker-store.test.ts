import { describe, it, before } from "node:test";
import assert from "node:assert/strict";

// Worker-store depends on window, document, and the state module's log().
// Set up minimal shims before importing.

type MessageHandler = (event: { source: unknown; data: unknown }) => void;
const messageHandlers: MessageHandler[] = [];

before(() => {
  const dataset: Record<string, string> = {};

  // Minimal document shim
  if (typeof globalThis.document === "undefined") {
    (globalThis as Record<string, unknown>).document = {
      documentElement: { dataset }
    };
  }

  // Minimal window shim — track addEventListener calls
  if (typeof globalThis.window === "undefined") {
    const windowShim = {
      addEventListener(eventType: string, handler: MessageHandler) {
        if (eventType === "message") {
          messageHandlers.push(handler);
        }
      }
    };
    (globalThis as Record<string, unknown>).window = windowShim;
  }
});

function dispatchBridgeMessage(type: string, payload: unknown): void {
  const event = {
    source: globalThis.window,
    data: {
      source: "teams-export-worker-hook",
      type,
      payload
    }
  };

  for (const handler of messageHandlers) {
    handler(event);
  }
}

// Dynamic import after shims are in place. We use a top-level await pattern
// via the test `before` hook to ensure proper sequencing.
let getWorkerMessage: typeof import("../../src/worker-store.js").getWorkerMessage;
let resolveUserId: typeof import("../../src/worker-store.js").resolveUserId;
let getWorkerStats: typeof import("../../src/worker-store.js").getWorkerStats;
let getCapturedToken: typeof import("../../src/worker-store.js").getCapturedToken;
let getCapturedRegion: typeof import("../../src/worker-store.js").getCapturedRegion;
let getCapturedTokenType: typeof import("../../src/worker-store.js").getCapturedTokenType;
let getLastSeenConversationId: typeof import("../../src/worker-store.js").getLastSeenConversationId;
let initWorkerStore: typeof import("../../src/worker-store.js").initWorkerStore;

before(async () => {
  const module = await import("../../src/worker-store.js");
  getWorkerMessage = module.getWorkerMessage;
  resolveUserId = module.resolveUserId;
  getWorkerStats = module.getWorkerStats;
  getCapturedToken = module.getCapturedToken;
  getCapturedRegion = module.getCapturedRegion;
  getCapturedTokenType = module.getCapturedTokenType;
  getLastSeenConversationId = module.getLastSeenConversationId;
  initWorkerStore = module.initWorkerStore;

  // Wire up the message listener
  initWorkerStore();
});

describe("worker-store initial state", () => {
  it("starts with empty stats", () => {
    const stats = getWorkerStats();
    assert.equal(stats.batches, 0);
    assert.equal(stats.messages, 0);
    assert.equal(stats.members, 0);
  });

  it("starts with no captured token", () => {
    assert.equal(getCapturedToken(), null);
  });

  it("starts with no captured region", () => {
    assert.equal(getCapturedRegion(), null);
  });

  it("starts with skypetoken as default token type", () => {
    assert.equal(getCapturedTokenType(), "skypetoken");
  });

  it("starts with no conversation ID", () => {
    assert.equal(getLastSeenConversationId(), null);
  });
});

describe("worker-store message-batch handling", () => {
  it("stores messages from a bridge event", () => {
    dispatchBridgeMessage("message-batch", {
      operationName: "TestQuery",
      requestId: "req-1",
      dataSource: "test",
      workerUrl: "https://example.com/worker.js",
      messages: [
        {
          id: "msg-100",
          clientMessageId: null,
          originalArrivalTime: "2026-03-16T10:00:00Z",
          author: "Alice",
          authorId: "user-alice",
          messageType: "Text",
          content: "Hello from Alice",
          subject: null,
          threadType: null,
          isDeleted: false,
          editedTime: null,
          emotionsSummary: null,
          emotions: null,
          diverseEmotions: null,
          quotedMessages: null,
          conversationId: "19:abc@thread.v2"
        },
        {
          id: "msg-101",
          clientMessageId: null,
          originalArrivalTime: "2026-03-16T10:01:00Z",
          author: "Bob",
          authorId: "user-bob",
          messageType: "Text",
          content: "Hello from Bob",
          subject: null,
          threadType: null,
          isDeleted: false,
          editedTime: null,
          emotionsSummary: null,
          emotions: null,
          diverseEmotions: null,
          quotedMessages: null,
          conversationId: "19:abc@thread.v2"
        }
      ]
    });

    const stats = getWorkerStats();
    assert.equal(stats.batches, 1);
    assert.equal(stats.messages, 2);
    assert.equal(stats.members, 2);
  });

  it("retrieves stored messages by ID", () => {
    const message = getWorkerMessage("msg-100");
    assert.ok(message);
    assert.equal(message!.author, "Alice");
    assert.equal(message!.content, "Hello from Alice");
  });

  it("returns undefined for unknown message IDs", () => {
    assert.equal(getWorkerMessage("nonexistent"), undefined);
  });

  it("normalizes message IDs with content- prefix", () => {
    const message = getWorkerMessage("content-msg-100");
    assert.ok(message);
    assert.equal(message!.author, "Alice");
  });

  it("normalizes message IDs with timestamp- prefix", () => {
    const message = getWorkerMessage("timestamp-msg-100");
    assert.ok(message);
    assert.equal(message!.author, "Alice");
  });

  it("records conversation ID from first message batch", () => {
    assert.equal(getLastSeenConversationId(), "19:abc@thread.v2");
  });

  it("resolves user IDs to display names", () => {
    assert.equal(resolveUserId("user-alice"), "Alice");
    assert.equal(resolveUserId("user-bob"), "Bob");
  });

  it("returns undefined for unknown user IDs", () => {
    assert.equal(resolveUserId("unknown-user"), undefined);
  });

  it("deduplicates messages on repeated batches", () => {
    const sizeBefore = getWorkerStats().messages;

    dispatchBridgeMessage("message-batch", {
      operationName: "TestQuery",
      requestId: "req-2",
      dataSource: "test",
      workerUrl: "https://example.com/worker.js",
      messages: [
        {
          id: "msg-100",
          clientMessageId: null,
          originalArrivalTime: "2026-03-16T10:00:00Z",
          author: "Alice",
          authorId: "user-alice",
          messageType: "Text",
          content: "Hello from Alice (updated)",
          subject: null,
          threadType: null,
          isDeleted: false,
          editedTime: null,
          emotionsSummary: null,
          emotions: null,
          diverseEmotions: null,
          quotedMessages: null,
          conversationId: "19:abc@thread.v2"
        }
      ]
    });

    // Message count should not increase for duplicate IDs
    assert.equal(getWorkerStats().messages, sizeBefore);
    // But the content should be updated
    assert.equal(getWorkerMessage("msg-100")!.content, "Hello from Alice (updated)");
  });

  it("ignores messages without an ID", () => {
    const sizeBefore = getWorkerStats().messages;

    dispatchBridgeMessage("message-batch", {
      operationName: "TestQuery",
      requestId: null,
      dataSource: null,
      workerUrl: "",
      messages: [
        {
          id: "",
          clientMessageId: null,
          originalArrivalTime: null,
          author: null,
          authorId: null,
          messageType: null,
          content: null,
          subject: null,
          threadType: null,
          isDeleted: false,
          editedTime: null,
          emotionsSummary: null,
          emotions: null,
          diverseEmotions: null,
          quotedMessages: null,
          conversationId: null
        }
      ]
    });

    assert.equal(getWorkerStats().messages, sizeBefore);
  });
});

describe("worker-store auth-token handling", () => {
  it("stores a bearer token from bridge event", () => {
    const token = "a".repeat(200); // Must be at least 100 chars
    dispatchBridgeMessage("auth-token", {
      token,
      region: "apac",
      tokenType: "bearer",
      source: "localStorage-msal"
    });

    assert.equal(getCapturedToken(), token);
    assert.equal(getCapturedRegion(), "apac");
    assert.equal(getCapturedTokenType(), "bearer");
  });

  it("ignores tokens shorter than 100 characters", () => {
    const previousToken = getCapturedToken();
    dispatchBridgeMessage("auth-token", {
      token: "short",
      region: "emea",
      tokenType: "bearer",
      source: "test"
    });

    assert.equal(getCapturedToken(), previousToken);
  });

  it("updates region when token is refreshed", () => {
    const token = "b".repeat(200);
    dispatchBridgeMessage("auth-token", {
      token,
      region: "emea",
      tokenType: "bearer",
      source: "test"
    });

    assert.equal(getCapturedRegion(), "emea");
  });
});

describe("worker-store members handling", () => {
  it("stores members from bridge event", () => {
    dispatchBridgeMessage("members", {
      members: [
        { id: "user-carol", displayName: "Carol" },
        { id: "user-dave", displayName: "Dave" }
      ]
    });

    assert.equal(resolveUserId("user-carol"), "Carol");
    assert.equal(resolveUserId("user-dave"), "Dave");
  });

  it("ignores members without an ID", () => {
    const membersBefore = getWorkerStats().members;
    dispatchBridgeMessage("members", {
      members: [
        { id: "", displayName: "Nobody" }
      ]
    });

    assert.equal(getWorkerStats().members, membersBefore);
  });

  it("does not overwrite existing member names", () => {
    dispatchBridgeMessage("members", {
      members: [
        { id: "user-alice", displayName: "Alice (Updated)" }
      ]
    });

    // registerMember skips if userId is already known
    assert.equal(resolveUserId("user-alice"), "Alice");
  });
});

describe("worker-store me identity handling", () => {
  it("registers me identity from bridge event", () => {
    dispatchBridgeMessage("me", {
      id: "user-me",
      displayName: "Current User"
    });

    assert.equal(resolveUserId("user-me"), "Current User");
  });
});

describe("worker-store bridge event filtering", () => {
  it("ignores events from other sources", () => {
    const statsBefore = getWorkerStats();

    for (const handler of messageHandlers) {
      handler({
        source: globalThis.window,
        data: {
          source: "some-other-extension",
          type: "message-batch",
          payload: {
            operationName: "Test",
            requestId: null,
            dataSource: null,
            workerUrl: "",
            messages: [{ id: "should-not-appear", author: "X", authorId: "x" }]
          }
        }
      });
    }

    assert.equal(getWorkerStats().batches, statsBefore.batches);
  });

  it("ignores events without a source field", () => {
    const statsBefore = getWorkerStats();

    for (const handler of messageHandlers) {
      handler({
        source: globalThis.window,
        data: { type: "message-batch", payload: {} }
      });
    }

    assert.equal(getWorkerStats().batches, statsBefore.batches);
  });

  it("ignores events with non-object data", () => {
    const statsBefore = getWorkerStats();

    for (const handler of messageHandlers) {
      handler({ source: globalThis.window, data: "string-data" });
    }

    assert.equal(getWorkerStats().batches, statsBefore.batches);
  });

  it("ignores events with null data", () => {
    const statsBefore = getWorkerStats();

    for (const handler of messageHandlers) {
      handler({ source: globalThis.window, data: null });
    }

    assert.equal(getWorkerStats().batches, statsBefore.batches);
  });
});
