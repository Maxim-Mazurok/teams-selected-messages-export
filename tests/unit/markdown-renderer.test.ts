import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  formatQuotedReplyLabel,
  renderQuotedReplyMarkdown,
  renderReactionsMarkdown
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
