import test from "node:test";
import assert from "node:assert/strict";
import {
  stripHtml,
  simplifyEmotionUsers,
  simplifyMessage,
  collectReactionUsers
} from "../../scripts/probe-worker-intercepts.mjs";

test("stripHtml removes markup and normalizes whitespace", () => {
  assert.equal(stripHtml("<p>Hello <strong>world</strong></p>"), "Hello world");
  assert.equal(stripHtml("<blockquote><p>Quoted</p></blockquote><p>Body</p>"), "Body");
});

test("simplifyEmotionUsers keeps keys, counts, and user ids", () => {
  assert.deepEqual(
    simplifyEmotionUsers([
      {
        key: "like",
        count: 2,
        users: [{ userId: "u1" }, { userId: "u2" }]
      }
    ]),
    [
      {
        key: "like",
        count: 2,
        userIds: ["u1", "u2"]
      }
    ]
  );
});

test("simplifyMessage keeps reaction metadata and content preview", () => {
  const message = simplifyMessage({
    id: "m1",
    originalArrivalTime: "2026-03-11T09:00:00.000Z",
    fromUser: { id: "u1", displayName: "Alice Example" },
    messageType: "RichText/Html",
    content: "<p>Hello <strong>world</strong></p>",
    emotionsSummary: [{ key: "like", count: 1 }],
    emotions: [{ key: "like", users: [{ userId: "u2" }] }],
    diverseEmotions: {
      items: [{ commonKey: "yes", users: [{ userId: "u2" }] }]
    }
  });

  assert.equal(message.author, "Alice Example");
  assert.equal(message.contentPreview, "Hello world");
  assert.deepEqual(message.emotionsSummary, [{ key: "like", count: 1 }]);
  assert.deepEqual(message.emotions, [{ key: "like", count: 1, userIds: ["u2"] }]);
  assert.deepEqual(message.diverseEmotions, [{ key: "yes", count: 1, userIds: ["u2"] }]);
});

test("collectReactionUsers de-duplicates per message, reaction, and user", () => {
  const results = collectReactionUsers([
    {
      id: "m1",
      emotions: [{ key: "like", userIds: ["u1", "u1"] }],
      diverseEmotions: [{ key: "yes", userIds: ["u2"] }]
    }
  ]);

  assert.deepEqual(results, [
    { messageId: "m1", reactionKey: "like", userId: "u1" },
    { messageId: "m1", reactionKey: "yes", userId: "u2" }
  ]);
});
