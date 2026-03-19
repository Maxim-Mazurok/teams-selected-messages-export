import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  formatQuotedReplyLabel,
  renderQuotedReplyMarkdown,
  renderReactionsMarkdown,
  renderMarkdown
} from "../../src/markdown-renderer.js";

describe("formatQuotedReplyLabel", () => {
  it("formats label with author and time", () => {
    assert.equal(
      formatQuotedReplyLabel({ author: "Alice", timeLabel: "2:30 PM", text: "hello" }),
      "Replying to Alice | 2:30 PM"
    );
  });

  it("formats label with author only", () => {
    assert.equal(
      formatQuotedReplyLabel({ author: "Alice", timeLabel: "", text: "hello" }),
      "Replying to Alice"
    );
  });

  it("formats label with time only", () => {
    assert.equal(
      formatQuotedReplyLabel({ author: "", timeLabel: "2:30 PM", text: "hello" }),
      "Replying to 2:30 PM"
    );
  });

  it("returns plain prefix when no author or time", () => {
    assert.equal(
      formatQuotedReplyLabel({ author: "", timeLabel: "", text: "hello" }),
      "Replying to"
    );
  });
});

describe("renderQuotedReplyMarkdown", () => {
  it("returns empty string when quote is null", () => {
    assert.equal(renderQuotedReplyMarkdown(null), "");
  });

  it("returns empty string when quote text is empty", () => {
    assert.equal(
      renderQuotedReplyMarkdown({ author: "Alice", timeLabel: "", text: "" }),
      ""
    );
  });

  it("renders quoted reply as blockquote", () => {
    const result = renderQuotedReplyMarkdown({
      author: "Alice",
      timeLabel: "2:30 PM",
      text: "Hello world"
    });

    assert.ok(result.includes("> Replying to Alice | 2:30 PM"));
    assert.ok(result.includes("> Hello world"));
  });

  it("handles multiline quoted text", () => {
    const result = renderQuotedReplyMarkdown({
      author: "Bob",
      timeLabel: "",
      text: "Line one\nLine two"
    });

    assert.ok(result.includes("> Line one"));
    assert.ok(result.includes("> Line two"));
  });
});

describe("renderReactionsMarkdown", () => {
  it("returns empty string for empty reactions", () => {
    assert.equal(renderReactionsMarkdown([]), "");
  });

  it("returns empty string for undefined reactions", () => {
    assert.equal(renderReactionsMarkdown(undefined), "");
  });

  it("renders reactions with actors", () => {
    const result = renderReactionsMarkdown([
      { emoji: "👍", name: "like", count: 2, actors: ["Alice", "Bob"] }
    ]);

    assert.equal(result, "Reactions: 👍 like by Alice, Bob");
  });

  it("renders reactions without actors using count", () => {
    const result = renderReactionsMarkdown([
      { emoji: "❤️", name: "heart", count: 3, actors: [] }
    ]);

    assert.equal(result, "Reactions: ❤️ heart x3");
  });

  it("renders multiple reactions separated by semicolons", () => {
    const result = renderReactionsMarkdown([
      { emoji: "👍", name: "like", count: 1, actors: ["Alice"] },
      { emoji: "😂", name: "laugh", count: 2, actors: [] }
    ]);

    assert.equal(result, "Reactions: 👍 like by Alice; 😂 laugh x2");
  });
});

describe("mention merging in markdown output", () => {
  it("merges consecutive mentions with capital first letters", () => {
    // This tests the regex pattern used in elementToMarkdown
    const markdown = "Hey @Jeremy @Kothe, check this out!";
    const merged = markdown
      .replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3")
      .replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)$/gm, "@$1 $2");
    assert.equal(merged, "Hey @Jeremy Kothe, check this out!");
  });

  it("merges multiple mention pairs in same text", () => {
    const markdown = "@Leo @Shchurov and @Alan @Markus";
    // Need to apply the regex repeatedly since the second pair has whitespace before @Alan
    let merged = markdown.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3");
    merged = merged.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)$/gm, "@$1 $2"); // For end-of-line cases
    assert.equal(merged, "@Leo Shchurov and @Alan Markus");
  });

  it("merges mentions at end of text", () => {
    const markdown = "Check with @Jim @Cooke";
    const merged = markdown
      .replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3")
      .replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)$/gm, "@$1 $2");
    assert.equal(merged, "Check with @Jim Cooke");
  });

  it("does not merge single mentions", () => {
    const markdown = "Hey @Maxim Mazurok!";
    const merged = markdown.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3");
    assert.equal(merged, "Hey @Maxim Mazurok!");
  });

  it("does not merge mentions without capital letters", () => {
    const markdown = "Contact @john @smith for details";
    const merged = markdown.replace(/@([A-Z][a-z]+)\s+@([A-Z][a-z]+)([\s,])/g, "@$1 $2$3");
    assert.equal(merged, "Contact @john @smith for details");
  });
});

describe("renderMarkdown thread reply formatting", () => {
  const baseMeta = {
    title: "Test Channel",
    sourceUrl: "https://teams.cloud.microsoft/",
    exportedAt: "2026-03-19",
    scope: "full-chat"
  };

  function makeMessage(overrides: Partial<import("../../src/types.js").MessageSnapshot> = {}): import("../../src/types.js").MessageSnapshot {
    return {
      id: "1",
      index: 0,
      author: "Alice",
      timeLabel: "2026-03-19 01:00",
      dateTime: "2026-03-19T01:00:00.000Z",
      subject: "",
      quote: null,
      reactions: [],
      html: "<p>Hello</p>",
      markdown: "Hello",
      plainText: "Hello",
      ...overrides
    };
  }

  it("renders root posts with h2 heading", () => {
    const result = renderMarkdown(
      [makeMessage({ isReply: false, threadId: "100" })],
      baseMeta
    );
    assert.ok(result.includes("## Alice | 2026-03-19 01:00"));
    assert.ok(!result.includes("### ↳"));
  });

  it("renders thread replies with h3 heading and arrow prefix", () => {
    const result = renderMarkdown(
      [makeMessage({ id: "200", isReply: true, threadId: "100", author: "Bob" })],
      baseMeta
    );
    assert.ok(result.includes("### ↳ Bob | 2026-03-19 01:00"));
    assert.ok(!result.includes("## Bob"));
  });

  it("renders messages without thread metadata as h2", () => {
    const result = renderMarkdown(
      [makeMessage()],
      baseMeta
    );
    assert.ok(result.includes("## Alice | 2026-03-19 01:00"));
    assert.ok(!result.includes("### ↳"));
  });
});
