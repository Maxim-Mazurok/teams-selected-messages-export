import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { getMessageRange, orderMessagesForExport } from "../../src/selection.js";
import type { MessageSnapshot } from "../../src/types.js";

describe("getMessageRange", () => {
  const messageIds = ["a", "b", "c", "d", "e"];

  it("returns range between anchor and target (forward)", () => {
    assert.deepEqual(getMessageRange(messageIds, "b", "d"), ["b", "c", "d"]);
  });

  it("returns range between anchor and target (backward)", () => {
    assert.deepEqual(getMessageRange(messageIds, "d", "b"), ["b", "c", "d"]);
  });

  it("returns single element when anchor equals target", () => {
    assert.deepEqual(getMessageRange(messageIds, "c", "c"), ["c"]);
  });

  it("returns only target when anchor is not found", () => {
    assert.deepEqual(getMessageRange(messageIds, "z", "c"), ["c"]);
  });

  it("returns only target when target is not found", () => {
    assert.deepEqual(getMessageRange(messageIds, "b", "z"), ["z"]);
  });

  it("returns full range for first to last", () => {
    assert.deepEqual(getMessageRange(messageIds, "a", "e"), [
      "a",
      "b",
      "c",
      "d",
      "e"
    ]);
  });
});

describe("orderMessagesForExport", () => {
  function createSnapshot(
    overrides: Partial<MessageSnapshot> & { id: string }
  ): MessageSnapshot {
    return {
      index: 0,
      author: "Test",
      timeLabel: "",
      dateTime: "",
      quote: null,
      reactions: [],
      html: "",
      markdown: "",
      plainText: "test",
      ...overrides
    };
  }

  it("sorts by dateTime when available", () => {
    const messages = [
      createSnapshot({ id: "2", dateTime: "2024-01-15T12:00:00Z" }),
      createSnapshot({ id: "1", dateTime: "2024-01-15T10:00:00Z" }),
      createSnapshot({ id: "3", dateTime: "2024-01-15T14:00:00Z" })
    ];

    const sorted = orderMessagesForExport(messages);
    assert.deepEqual(
      sorted.map((message) => message.id),
      ["1", "2", "3"]
    );
  });

  it("falls back to numeric id when dates are equal", () => {
    const messages = [
      createSnapshot({ id: "300", dateTime: "2024-01-15T12:00:00Z" }),
      createSnapshot({ id: "100", dateTime: "2024-01-15T12:00:00Z" }),
      createSnapshot({ id: "200", dateTime: "2024-01-15T12:00:00Z" })
    ];

    const sorted = orderMessagesForExport(messages);
    assert.deepEqual(
      sorted.map((message) => message.id),
      ["100", "200", "300"]
    );
  });

  it("falls back to captureOrder when no dates or numeric ids", () => {
    const messages = [
      createSnapshot({ id: "msg-c", captureOrder: 2 }),
      createSnapshot({ id: "msg-a", captureOrder: 0 }),
      createSnapshot({ id: "msg-b", captureOrder: 1 })
    ];

    const sorted = orderMessagesForExport(messages);
    assert.deepEqual(
      sorted.map((message) => message.id),
      ["msg-a", "msg-b", "msg-c"]
    );
  });

  it("does not mutate the original array", () => {
    const messages = [
      createSnapshot({ id: "2", dateTime: "2024-01-15T12:00:00Z" }),
      createSnapshot({ id: "1", dateTime: "2024-01-15T10:00:00Z" })
    ];

    const original = [...messages];
    orderMessagesForExport(messages);
    assert.deepEqual(
      messages.map((message) => message.id),
      original.map((message) => message.id)
    );
  });
});
