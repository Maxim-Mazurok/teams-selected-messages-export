import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { parseChannelTimeAriaLabel } from "../../src/content-extraction.js";

describe("parseChannelTimeAriaLabel", () => {
  it("parses full date string from aria-label", () => {
    const result = parseChannelTimeAriaLabel("09 March 2026 14:42");
    assert.ok(result, "Should return a non-empty string");
    const date = new Date(result);
    assert.equal(date.getFullYear(), 2026);
    assert.equal(date.getMonth(), 2); // March is month 2 (0-indexed)
    assert.equal(date.getDate(), 9);
  });

  it("returns empty string for empty input", () => {
    assert.equal(parseChannelTimeAriaLabel(""), "");
  });

  it("returns empty string for whitespace-only input", () => {
    assert.equal(parseChannelTimeAriaLabel("   "), "");
  });

  it("returns empty string for unparseable date string", () => {
    assert.equal(parseChannelTimeAriaLabel("Monday 14:42"), "");
  });

  it("returns ISO string format", () => {
    const result = parseChannelTimeAriaLabel("10 March 2026 09:30");
    assert.match(result, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}.\d{3}Z$/);
  });

  it("parses various date formats", () => {
    const result = parseChannelTimeAriaLabel("March 10, 2026 11:15");
    assert.ok(result, "Should parse US-style date format");
    const date = new Date(result);
    assert.equal(date.getFullYear(), 2026);
  });
});
