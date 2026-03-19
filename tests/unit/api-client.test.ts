import { describe, it, before } from "node:test";
import assert from "node:assert/strict";

// convertApiMessagesToSnapshots calls stripHtmlTags which uses document.createElement.
// Provide a minimal shim so the tests can run in Node.js without a full DOM.
before(() => {
  if (typeof globalThis.document === "undefined") {
    const elements = new Map<string, { innerHTML: string; textContent: string | null }>();

    (globalThis as Record<string, unknown>).document = {
      createElement(tag: string) {
        const element = { innerHTML: "", textContent: null as string | null };
        Object.defineProperty(element, "textContent", {
          get() {
            // Primitive HTML tag stripping — sufficient for test data
            return (element.innerHTML || "")
              .replace(/<[^>]*>/g, "")
              .replace(/&amp;/g, "&")
              .replace(/&lt;/g, "<")
              .replace(/&gt;/g, ">")
              .replace(/&quot;/g, '"')
              .trim();
          }
        });
        elements.set(tag, element);
        return element;
      },
      documentElement: { dataset: {} }
    };
  }
});

import {
  extractConversationIdFromUrl,
  convertApiMessagesToSnapshots
} from "../../src/api-client.js";

// ── extractConversationIdFromUrl ──────────────────────────────────────────────

describe("extractConversationIdFromUrl", () => {
  const originalLocation = globalThis.location;

  function setLocation(url: string): void {
    const parsed = new URL(url);
    Object.defineProperty(globalThis, "location", {
      value: {
        pathname: parsed.pathname,
        hash: parsed.hash,
        href: parsed.href
      },
      writable: true,
      configurable: true
    });
  }

  function restoreLocation(): void {
    Object.defineProperty(globalThis, "location", {
      value: originalLocation,
      writable: true,
      configurable: true
    });
  }

  it("extracts conversation ID from /conversations/ path", () => {
    setLocation("https://teams.microsoft.com/conversations/19:abc123@thread.v2");
    try {
      assert.equal(
        extractConversationIdFromUrl(),
        "19:abc123@thread.v2"
      );
    } finally {
      restoreLocation();
    }
  });

  it("extracts conversation ID from /l/message/ path", () => {
    setLocation("https://teams.microsoft.com/l/message/19:abc123@thread.v2/details");
    try {
      assert.equal(
        extractConversationIdFromUrl(),
        "19:abc123@thread.v2"
      );
    } finally {
      restoreLocation();
    }
  });

  it("extracts conversation ID from /l/chat/ path", () => {
    setLocation("https://teams.microsoft.com/l/chat/19:abc123@thread.v2/details");
    try {
      assert.equal(
        extractConversationIdFromUrl(),
        "19:abc123@thread.v2"
      );
    } finally {
      restoreLocation();
    }
  });

  it("extracts conversation ID from hash fragment", () => {
    setLocation("https://teams.microsoft.com/#/conversations/19:abc123@thread.v2");
    try {
      assert.equal(
        extractConversationIdFromUrl(),
        "19:abc123@thread.v2"
      );
    } finally {
      restoreLocation();
    }
  });

  it("extracts conversation ID from hash-based /l/message/ path", () => {
    setLocation("https://teams.microsoft.com/#/l/message/19:abc123@thread.v2/details");
    try {
      assert.equal(
        extractConversationIdFromUrl(),
        "19:abc123@thread.v2"
      );
    } finally {
      restoreLocation();
    }
  });

  it("returns null for Teams Cloud root URL (no path)", () => {
    setLocation("https://teams.cloud.microsoft/");
    try {
      assert.equal(extractConversationIdFromUrl(), null);
    } finally {
      restoreLocation();
    }
  });

  it("returns null for URLs without a conversation segment", () => {
    setLocation("https://teams.microsoft.com/v2/dashboard");
    try {
      assert.equal(extractConversationIdFromUrl(), null);
    } finally {
      restoreLocation();
    }
  });

  it("decodes encoded conversation IDs", () => {
    setLocation("https://teams.microsoft.com/conversations/19%3Aabc123%40thread.v2");
    try {
      assert.equal(
        extractConversationIdFromUrl(),
        "19:abc123@thread.v2"
      );
    } finally {
      restoreLocation();
    }
  });
});

// ── convertApiMessagesToSnapshots ─────────────────────────────────────────────

describe("convertApiMessagesToSnapshots", () => {
  it("converts a single text message", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1001",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Hello world",
        originalarrivaltime: "2026-03-16T10:00:00.000Z"
      }
    ]);

    assert.equal(snapshots.length, 1);
    assert.equal(snapshots[0].id, "1001");
    assert.equal(snapshots[0].author, "Alice");
    assert.equal(snapshots[0].plainText, "Hello world");
    assert.equal(snapshots[0].markdown, "Hello world");
    assert.equal(snapshots[0].html, "<p>Hello world</p>");
  });

  it("converts RichText/Html messages by stripping tags for plainText and markdown", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1002",
        type: "Message",
        messagetype: "RichText/Html",
        imdisplayname: "Bob",
        content: "<p>Hello <b>world</b></p>",
        originalarrivaltime: "2026-03-16T10:01:00.000Z"
      }
    ]);

    assert.equal(snapshots.length, 1);
    assert.equal(snapshots[0].plainText, "Hello world");
    assert.equal(snapshots[0].markdown, "Hello world");
    assert.equal(snapshots[0].html, "<p>Hello <b>world</b></p>");
  });

  it("filters out system/control messages", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1003",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Real message"
      },
      {
        id: "1004",
        type: "Message",
        messagetype: "ThreadActivity/AddMember",
        content: "Alice added Bob"
      },
      {
        id: "1005",
        type: "Event/Call",
        content: "Call started"
      },
      {
        id: "1006",
        type: "Message",
        messagetype: "MessageDelete",
        content: ""
      }
    ]);

    assert.equal(snapshots.length, 1);
    assert.equal(snapshots[0].id, "1003");
  });

  it("reverses API order to chronological", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "2",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Bob",
        content: "Second message",
        originalarrivaltime: "2026-03-16T10:01:00.000Z"
      },
      {
        id: "1",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "First message",
        originalarrivaltime: "2026-03-16T10:00:00.000Z"
      }
    ]);

    assert.equal(snapshots.length, 2);
    assert.equal(snapshots[0].id, "1");
    assert.equal(snapshots[0].author, "Alice");
    assert.equal(snapshots[1].id, "2");
    assert.equal(snapshots[1].author, "Bob");
  });

  it("converts reactions/emotions to ReactionInfo", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1010",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Great news!",
        emotions: [
          {
            key: "like",
            users: [
              { mri: "8:orgid:user1", displayName: "Bob", time: "2026-03-16T10:05:00Z" },
              { mri: "8:orgid:user2", displayName: "Carol", time: "2026-03-16T10:06:00Z" }
            ]
          },
          {
            key: "heart",
            users: [
              { mri: "8:orgid:user1", displayName: "Bob", time: "2026-03-16T10:07:00Z" }
            ]
          }
        ]
      }
    ]);

    assert.equal(snapshots.length, 1);
    assert.equal(snapshots[0].reactions.length, 2);
    assert.equal(snapshots[0].reactions[0].emoji, "like");
    assert.equal(snapshots[0].reactions[0].name, "like");
    assert.equal(snapshots[0].reactions[0].count, 2);
    assert.deepEqual(snapshots[0].reactions[0].actors, ["Bob", "Carol"]);
    assert.equal(snapshots[0].reactions[1].emoji, "heart");
    assert.equal(snapshots[0].reactions[1].count, 1);
  });

  it("skips emotions with no users", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1011",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Test",
        emotions: [
          { key: "like", users: [] },
          { key: "heart", users: [{ mri: "8:orgid:user1", displayName: "Bob", time: "t" }] }
        ]
      }
    ]);

    assert.equal(snapshots[0].reactions.length, 1);
    assert.equal(snapshots[0].reactions[0].emoji, "heart");
  });

  it("handles missing emotions gracefully", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1012",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "No reactions"
      }
    ]);

    assert.deepEqual(snapshots[0].reactions, []);
  });

  it("uses imdisplayname for author, falls back to Unknown author", () => {
    const withAuthor = convertApiMessagesToSnapshots([
      { id: "a", type: "Message", messagetype: "Text", imdisplayname: "Alice", content: "Hi" }
    ]);
    const withoutAuthor = convertApiMessagesToSnapshots([
      { id: "b", type: "Message", messagetype: "Text", content: "Hi" }
    ]);

    assert.equal(withAuthor[0].author, "Alice");
    assert.equal(withoutAuthor[0].author, "Unknown author");
  });

  it("extracts subject from properties", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1020",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Post body",
        properties: { subject: "Important topic" }
      }
    ]);

    assert.equal(snapshots[0].subject, "Important topic");
  });

  it("handles empty properties gracefully", () => {
    const snapshots = convertApiMessagesToSnapshots([
      { id: "1021", type: "Message", messagetype: "Text", imdisplayname: "A", content: "X" }
    ]);
    assert.equal(snapshots[0].subject, "");
  });

  it("formats timestamps in ISO format for dateTime", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1030",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Test",
        originalarrivaltime: "2026-03-16T10:30:00.000Z"
      }
    ]);

    assert.equal(snapshots[0].dateTime, "2026-03-16T10:30:00.000Z");
    assert.ok(snapshots[0].timeLabel.length > 0, "timeLabel should not be empty");
  });

  it("prefers originalarrivaltime over composetime", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1031",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Test",
        originalarrivaltime: "2026-03-16T10:30:00.000Z",
        composetime: "2026-03-16T10:29:00.000Z"
      }
    ]);

    assert.equal(snapshots[0].dateTime, "2026-03-16T10:30:00.000Z");
  });

  it("falls back to composetime when originalarrivaltime is missing", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1032",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Test",
        composetime: "2026-03-16T10:29:00.000Z"
      }
    ]);

    assert.equal(snapshots[0].dateTime, "2026-03-16T10:29:00.000Z");
  });

  it("handles missing timestamps gracefully", () => {
    const snapshots = convertApiMessagesToSnapshots([
      { id: "1033", type: "Message", messagetype: "Text", imdisplayname: "A", content: "X" }
    ]);

    assert.equal(snapshots[0].dateTime, "");
    assert.equal(snapshots[0].timeLabel, "");
  });

  it("returns empty array for empty input", () => {
    assert.deepEqual(convertApiMessagesToSnapshots([]), []);
  });

  it("assigns sequential index and captureOrder", () => {
    const snapshots = convertApiMessagesToSnapshots([
      { id: "3", type: "Message", messagetype: "Text", imdisplayname: "C", content: "Third" },
      { id: "2", type: "Message", messagetype: "Text", imdisplayname: "B", content: "Second" },
      { id: "1", type: "Message", messagetype: "Text", imdisplayname: "A", content: "First" }
    ]);

    // After reversal: 1, 2, 3
    assert.equal(snapshots[0].index, 0);
    assert.equal(snapshots[0].captureOrder, 0);
    assert.equal(snapshots[1].index, 1);
    assert.equal(snapshots[1].captureOrder, 1);
    assert.equal(snapshots[2].index, 2);
    assert.equal(snapshots[2].captureOrder, 2);
  });

  it("sets quote to null for all API messages", () => {
    const snapshots = convertApiMessagesToSnapshots([
      { id: "1040", type: "Message", messagetype: "Text", imdisplayname: "A", content: "X" }
    ]);

    assert.equal(snapshots[0].quote, null);
  });

  it("handles content with HTML-like text in plain Text messages", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1050",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Check out <this> approach"
      }
    ]);

    // Text messages with < are treated as containing HTML
    assert.equal(snapshots[0].plainText, "Check out  approach");
  });

  it("handles missing content gracefully", () => {
    const snapshots = convertApiMessagesToSnapshots([
      { id: "1060", type: "Message", messagetype: "Text", imdisplayname: "Alice" }
    ]);

    assert.equal(snapshots[0].plainText, "");
    assert.equal(snapshots[0].markdown, "");
  });

  it("filters all ThreadActivity subtypes", () => {
    const activityTypes = [
      "ThreadActivity/AddMember",
      "ThreadActivity/MemberJoined",
      "ThreadActivity/MemberLeft",
      "ThreadActivity/DeleteMember",
      "ThreadActivity/TopicUpdate",
      "ThreadActivity/HistoryDisclosedUpdate"
    ];

    for (const activityType of activityTypes) {
      const snapshots = convertApiMessagesToSnapshots([
        { id: "x", type: "Message", messagetype: activityType, content: "system" }
      ]);
      assert.equal(snapshots.length, 0, `Should filter ${activityType}`);
    }
  });

  it("uses MRI as actor fallback when displayName is missing", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1070",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Test",
        emotions: [
          {
            key: "like",
            users: [
              { mri: "8:orgid:user-no-name", time: "t" }
            ]
          }
        ]
      }
    ]);

    assert.equal(snapshots[0].reactions[0].actors[0], "8:orgid:user-no-name");
  });

  // ── Channel thread grouping ───────────────────────────────────────────────

  it("groups channel messages by thread with root posts before replies", () => {
    const snapshots = convertApiMessagesToSnapshots([
      // Newest first (API order) — reply to thread 100, then root of thread 200, then root of thread 100
      {
        id: "300",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Charlie",
        content: "Reply to thread 100",
        composetime: "2026-03-19T03:00:00.000Z",
        conversationLink: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:abc@thread.tacv2;messageid=100"
      },
      {
        id: "200",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Bob",
        content: "Root of thread 200",
        composetime: "2026-03-19T02:00:00.000Z",
        conversationLink: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:abc@thread.tacv2;messageid=200"
      },
      {
        id: "100",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Root of thread 100",
        composetime: "2026-03-19T01:00:00.000Z",
        conversationLink: "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:abc@thread.tacv2;messageid=100"
      }
    ]);

    // Thread 100 (root) should come first (oldest root), then its reply, then thread 200
    assert.equal(snapshots.length, 3);
    assert.equal(snapshots[0].id, "100"); // root of thread 100
    assert.equal(snapshots[0].isReply, false);
    assert.equal(snapshots[0].threadId, "100");
    assert.equal(snapshots[1].id, "300"); // reply to thread 100
    assert.equal(snapshots[1].isReply, true);
    assert.equal(snapshots[1].threadId, "100");
    assert.equal(snapshots[2].id, "200"); // root of thread 200
    assert.equal(snapshots[2].isReply, false);
    assert.equal(snapshots[2].threadId, "200");
  });

  it("does not set threadId or isReply for DM messages without conversationLink", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "1",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Hello",
        composetime: "2026-03-19T01:00:00.000Z"
      },
      {
        id: "2",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Bob",
        content: "Hi",
        composetime: "2026-03-19T02:00:00.000Z"
      }
    ]);

    // DM messages: chronological order, no thread metadata
    assert.equal(snapshots.length, 2);
    assert.equal(snapshots[0].id, "2"); // reversed to chronological (newest-first → oldest-first)
    assert.equal(snapshots[0].threadId, undefined);
    assert.equal(snapshots[0].isReply, undefined);
  });

  it("sorts replies within a thread chronologically", () => {
    const snapshots = convertApiMessagesToSnapshots([
      // API returns newest first
      {
        id: "400",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Dave",
        content: "Late reply",
        composetime: "2026-03-19T04:00:00.000Z",
        conversationLink: "https://x.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:ch@thread.tacv2;messageid=100"
      },
      {
        id: "200",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Bob",
        content: "Early reply",
        composetime: "2026-03-19T02:00:00.000Z",
        conversationLink: "https://x.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:ch@thread.tacv2;messageid=100"
      },
      {
        id: "100",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Root post",
        composetime: "2026-03-19T01:00:00.000Z",
        conversationLink: "https://x.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:ch@thread.tacv2;messageid=100"
      }
    ]);

    assert.equal(snapshots[0].id, "100"); // root
    assert.equal(snapshots[1].id, "200"); // early reply
    assert.equal(snapshots[2].id, "400"); // late reply
  });

  it("handles channel messages where root post is missing from response", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "300",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Charlie",
        content: "Reply without root",
        composetime: "2026-03-19T03:00:00.000Z",
        conversationLink: "https://x.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:ch@thread.tacv2;messageid=100"
      }
    ]);

    // Should still work — reply is present even though root isn't
    assert.equal(snapshots.length, 1);
    assert.equal(snapshots[0].id, "300");
    assert.equal(snapshots[0].isReply, true);
    assert.equal(snapshots[0].threadId, "100");
  });

  it("assigns sequential indexes after thread grouping", () => {
    const snapshots = convertApiMessagesToSnapshots([
      {
        id: "300",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Charlie",
        content: "Reply",
        composetime: "2026-03-19T03:00:00.000Z",
        conversationLink: "https://x.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:ch@thread.tacv2;messageid=100"
      },
      {
        id: "100",
        type: "Message",
        messagetype: "Text",
        imdisplayname: "Alice",
        content: "Root",
        composetime: "2026-03-19T01:00:00.000Z",
        conversationLink: "https://x.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19:ch@thread.tacv2;messageid=100"
      }
    ]);

    assert.equal(snapshots[0].index, 0);
    assert.equal(snapshots[0].captureOrder, 0);
    assert.equal(snapshots[1].index, 1);
    assert.equal(snapshots[1].captureOrder, 1);
  });
});
