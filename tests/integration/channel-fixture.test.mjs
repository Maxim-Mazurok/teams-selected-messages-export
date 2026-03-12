import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import puppeteer from "puppeteer-core";
import { findChromeExecutable } from "../helpers/chrome-path.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "../..");
const sourcePath = path.join(rootDir, "dist", "content-script.js");
const channelFixturePath = path.join(rootDir, "tests", "fixtures", "teams-channel-fixture.html");

async function injectPrototype(page) {
  const source = await fs.readFile(sourcePath, "utf8");
  await page.evaluate((script) => {
    eval(script);
  }, source);
  await page.waitForFunction(() => window.__teamsMessageExporter?.state?.messages?.length >= 1, {
    timeout: 10000
  });
}

test("channel fixture detects individual posts and replies as separate selectable messages", async () => {
  const executablePath = await findChromeExecutable();
  const browser = await puppeteer.launch({
    executablePath,
    headless: "new",
    args: ["--no-first-run", "--no-default-browser-check", "--disable-extensions"]
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1440, height: 1200 });
    await page.evaluateOnNewDocument(() => {
      Object.defineProperty(navigator, "clipboard", {
        configurable: true,
        value: {
          async writeText(text) {
            window.__copiedText = text;
          }
        }
      });
    });

    await page.goto(pathToFileURL(channelFixturePath).href, { waitUntil: "load" });
    await injectPrototype(page);
    await page.waitForFunction(() => document.querySelector(".tsm-export__dock"));

    const messageCount = await page.evaluate(() => window.__teamsMessageExporter.state.messages.length);
    assert.equal(messageCount, 5, "Should detect 3 posts + 2 replies = 5 individual messages");

    const strategyName = await page.evaluate(() => window.__teamsMessageExporter.state.strategy?.name || "");
    assert.equal(strategyName, "teams-channel-pane-message", "Should use channel strategy");

    const fullHistoryResult = await page.evaluate(async () => {
      const result = await window.__teamsMessageExporter.exportFullHistory("md", {
        download: false,
        closePanel: false
      });

      return {
        count: result?.count ?? 0,
        content: result?.content ?? ""
      };
    });

    assert.equal(fullHistoryResult.count, 5, "Should export 5 individual messages");

    assert.match(fullHistoryResult.content, /Alice Example/, "Should contain post author Alice");
    assert.match(fullHistoryResult.content, /Charlie Example/, "Should contain post author Charlie");
    assert.match(fullHistoryResult.content, /Dana Example/, "Should contain post author Dana");
    assert.match(fullHistoryResult.content, /Bob Example/, "Should contain reply author Bob");

    assert.match(fullHistoryResult.content, /\*\*AI Discussion Topic\*\*/, "Should contain subject line rendered bold");

    assert.match(fullHistoryResult.content, /GPT-5 API/, "Should contain post body text");
    assert.match(fullHistoryResult.content, /function calling is much improved/, "Should contain reply body from Bob");
    assert.match(fullHistoryResult.content, /standup at 10am/, "Should contain Charlie's post body");

    assert.match(fullHistoryResult.content, /@Alice Example/, "Should contain mention in Dana's post");
    assert.match(fullHistoryResult.content, /benchmark results/, "Should contain Dana's post body");
    assert.match(fullHistoryResult.content, /latency dropped by 40%/, "Should contain Alice's reply in thread 3");

    assert.match(fullHistoryResult.content, /Reactions:.*👍.*like/, "Should contain like reaction on post");
    assert.match(fullHistoryResult.content, /Reactions:.*❤️.*heart/, "Should contain heart reaction on reply");

    const fullHistoryHtmlResult = await page.evaluate(async () => {
      const result = await window.__teamsMessageExporter.exportFullHistory("html", {
        download: false,
        closePanel: false
      });

      return {
        count: result?.count ?? 0,
        content: result?.content ?? ""
      };
    });

    assert.equal(fullHistoryHtmlResult.count, 5);
    assert.match(fullHistoryHtmlResult.content, /<!doctype html>/i);
    assert.match(fullHistoryHtmlResult.content, /AI Technologies - Teams export/);
    assert.match(fullHistoryHtmlResult.content, /<h3>AI Discussion Topic<\/h3>/);

    await page.click(".tsm-export__launcher");
    await page.waitForFunction(() => document.body.classList.contains("tsm-export--selection-active"));

    const selectableMessages = await page.evaluate(
      () => Array.from(document.querySelectorAll('.tsm-export__message[data-tsm-message-id]')).length
    );
    assert.equal(selectableMessages, 5, "Should have 5 selectable message elements");

    const firstMessageId = await page.evaluate(
      () => document.querySelector('.tsm-export__message[data-tsm-message-id]')?.dataset.tsmMessageId || ""
    );
    assert.ok(firstMessageId, "First post should have a message ID");

    await page.click(`[data-tsm-message-id="${firstMessageId}"]`, { offset: { x: 40, y: 20 } });
    await page.waitForFunction(() => window.__teamsMessageExporter.state.selectedIds.size === 1);

    await page.click('.tsm-export__panel [data-action="copy-md"]');
    await page.waitForFunction(
      () => !document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open")
    );

    const copiedText = await page.evaluate(() => window.__copiedText || "");
    assert.match(copiedText, /Alice Example/, "Copied post should contain the post author");
    assert.match(copiedText, /GPT-5 API/, "Copied post should contain the post body");
    assert.doesNotMatch(copiedText, /function calling is much improved/, "Copied post should NOT contain reply body (replies are separate)");
  } finally {
    await browser.close();
  }
});
