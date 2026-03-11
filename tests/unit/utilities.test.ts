import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  normalizeText,
  escapeHtml,
  filenameSafe,
  parseRgbColor,
  relativeLuminanceChannel,
  isDarkRgb,
  inlineMarkdown,
  toComparableTime,
  toComparableNumericId
} from "../../src/utilities.js";

describe("normalizeText", () => {
  it("collapses multiple spaces into one", () => {
    assert.equal(normalizeText("hello   world"), "hello world");
  });

  it("trims leading and trailing whitespace", () => {
    assert.equal(normalizeText("  hello  "), "hello");
  });

  it("collapses tabs and newlines", () => {
    assert.equal(normalizeText("hello\n\tworld"), "hello world");
  });

  it("returns empty string for blank input", () => {
    assert.equal(normalizeText("   "), "");
  });
});

describe("escapeHtml", () => {
  it("escapes ampersands", () => {
    assert.equal(escapeHtml("a & b"), "a &amp; b");
  });

  it("escapes angle brackets", () => {
    assert.equal(escapeHtml("<div>"), "&lt;div&gt;");
  });

  it("escapes double quotes", () => {
    assert.equal(escapeHtml('a "b" c'), "a &quot;b&quot; c");
  });

  it("handles strings with multiple special characters", () => {
    assert.equal(
      escapeHtml('<a href="x">&</a>'),
      "&lt;a href=&quot;x&quot;&gt;&amp;&lt;/a&gt;"
    );
  });

  it("returns empty string for empty input", () => {
    assert.equal(escapeHtml(""), "");
  });
});

describe("filenameSafe", () => {
  it("replaces special characters with hyphens and lowercases", () => {
    assert.equal(filenameSafe("Hello World!"), "hello-world");
  });

  it("collapses multiple hyphens", () => {
    assert.equal(filenameSafe("a---b"), "a-b");
  });

  it("removes leading and trailing hyphens", () => {
    assert.equal(filenameSafe("--hello--"), "hello");
  });

  it("preserves dots and underscores", () => {
    assert.equal(filenameSafe("file_name.txt"), "file_name.txt");
  });
});

describe("parseRgbColor", () => {
  it("parses rgb color string", () => {
    const color = parseRgbColor("rgb(10, 20, 30)");
    assert.deepEqual(color, { red: 10, green: 20, blue: 30, alpha: 1 });
  });

  it("parses rgba color string", () => {
    const color = parseRgbColor("rgba(10, 20, 30, 0.5)");
    assert.deepEqual(color, { red: 10, green: 20, blue: 30, alpha: 0.5 });
  });

  it("returns null for non-rgb strings", () => {
    assert.equal(parseRgbColor("#fff"), null);
  });

  it("returns null for null input", () => {
    assert.equal(parseRgbColor(null), null);
  });

  it("returns null for empty string", () => {
    assert.equal(parseRgbColor(""), null);
  });
});

describe("relativeLuminanceChannel", () => {
  it("returns 0 for channel value 0", () => {
    assert.equal(relativeLuminanceChannel(0), 0);
  });

  it("returns 1 for channel value 255", () => {
    assert.ok(Math.abs(relativeLuminanceChannel(255) - 1) < 0.001);
  });

  it("uses linear formula for low values", () => {
    const result = relativeLuminanceChannel(5);
    const expected = 5 / 255 / 12.92;
    assert.ok(Math.abs(result - expected) < 0.0001);
  });
});

describe("isDarkRgb", () => {
  it("returns true for black", () => {
    assert.equal(isDarkRgb({ red: 0, green: 0, blue: 0, alpha: 1 }), true);
  });

  it("returns false for white", () => {
    assert.equal(isDarkRgb({ red: 255, green: 255, blue: 255, alpha: 1 }), false);
  });

  it("returns null for transparent", () => {
    assert.equal(isDarkRgb({ red: 0, green: 0, blue: 0, alpha: 0 }), null);
  });

  it("returns null for null input", () => {
    assert.equal(isDarkRgb(null), null);
  });
});

describe("inlineMarkdown", () => {
  it("trims and collapses whitespace", () => {
    assert.equal(inlineMarkdown("  hello   world  "), "hello world");
  });

  it("returns activeHref when text is empty", () => {
    assert.equal(inlineMarkdown("   ", "https://example.com"), "https://example.com");
  });

  it("returns empty string when both are empty", () => {
    assert.equal(inlineMarkdown(""), "");
  });
});

describe("toComparableTime", () => {
  it("parses valid ISO date string", () => {
    const result = toComparableTime("2024-01-15T10:30:00Z");
    assert.equal(typeof result, "number");
    assert.ok(result! > 0);
  });

  it("returns null for invalid date string", () => {
    assert.equal(toComparableTime("not-a-date"), null);
  });

  it("returns null for undefined", () => {
    assert.equal(toComparableTime(undefined), null);
  });
});

describe("toComparableNumericId", () => {
  it("parses numeric string", () => {
    assert.equal(toComparableNumericId("12345"), 12345);
  });

  it("returns null for non-numeric string", () => {
    assert.equal(toComparableNumericId("abc123"), null);
  });

  it("returns null for empty string", () => {
    assert.equal(toComparableNumericId(""), null);
  });

  it("returns null for undefined", () => {
    assert.equal(toComparableNumericId(undefined), null);
  });

  it("returns null for string with mixed content", () => {
    assert.equal(toComparableNumericId("123abc"), null);
  });
});
