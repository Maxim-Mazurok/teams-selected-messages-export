import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  getExportMessageCountLabel,
  getExportMessageSummary,
  getExportScopeSuffix
} from "../../src/export-helpers.js";

describe("getExportMessageCountLabel", () => {
  it("returns full chat label for full-chat scope", () => {
    assert.equal(
      getExportMessageCountLabel(42, "full-chat"),
      "Full chat history messages: 42"
    );
  });

  it("returns selection label for selection scope", () => {
    assert.equal(
      getExportMessageCountLabel(5, "selection"),
      "Selected messages: 5"
    );
  });

  it("returns selection label for unknown scope", () => {
    assert.equal(
      getExportMessageCountLabel(1, "other"),
      "Selected messages: 1"
    );
  });
});

describe("getExportMessageSummary", () => {
  it("returns full chat summary for full-chat scope", () => {
    assert.equal(
      getExportMessageSummary(10, "full-chat"),
      "10 messages captured from the full chat history."
    );
  });

  it("uses singular when count is 1", () => {
    assert.equal(
      getExportMessageSummary(1, "full-chat"),
      "1 message captured from the full chat history."
    );
  });

  it("returns selection summary for selection scope", () => {
    assert.equal(
      getExportMessageSummary(3, "selection"),
      "3 messages selected."
    );
  });

  it("returns singular selection summary", () => {
    assert.equal(
      getExportMessageSummary(1, "selection"),
      "1 message selected."
    );
  });
});

describe("getExportScopeSuffix", () => {
  it("returns full-chat suffix", () => {
    assert.equal(getExportScopeSuffix("full-chat"), "-full-chat");
  });

  it("returns empty string for selection scope", () => {
    assert.equal(getExportScopeSuffix("selection"), "");
  });

  it("returns empty string for unknown scope", () => {
    assert.equal(getExportScopeSuffix("anything"), "");
  });
});
