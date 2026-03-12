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

test("channel fixture detects channel threads and exports with reply content", async () => {
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
    assert.equal(messageCount, 3, "Should detect 3 channel threads");

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

    assert.equal(fullHistoryResult.count, 3, "Should export 3 threads");

    assert.match(fullHistoryResult.content, /Alice Example/, "Should contain post author Alice");
    assert.match(fullHistoryResult.content, /Charlie Example/, "Should contain post author Charlie");
    assert.match(fullHistoryResult.content, /Dana Example/, "Should contain post author Dana");

    assert.match(fullHistoryResult.content, /\*\*AI Discussion Topic\*\*/, "Should contain subject line rendered bold");

    assert.match(fullHistoryResult.content, /GPT-5 API/, "Should contain post body text");
    assert.match(fullHistoryResult.content, /function calling is much improved/, "Should contain reply body text from thread 1");

    assert.match(fullHistoryResult.content, /Bob Example/, "Reply author Bob should appear in content");

    assert.match(fullHistoryResult.content, /standup at 10am/, "Should contain Charlie's post body");

    assert.match(fullHistoryResult.content, /@Alice Example/, "Should contain mention in Dana's post");
    assert.match(fullHistoryResult.content, /benchmark results/, "Should contain Dana's post body");
    assert.match(fullHistoryResult.content, /latency dropped by 40%/, "Should contain Alice's reply in thread 3");

    assert.match(fullHistoryResult.content, /Reactions:.*👍.*like/, "Should contain like reaction");
    assert.match(fullHistoryResult.content, /Reactions:.*❤️.*heart/, "Should contain heart reaction");

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

    assert.equal(fullHistoryHtmlResult.count, 3);
    assert.match(fullHistoryHtmlResult.content, /<!doctype html>/i);
    assert.match(fullHistoryHtmlResult.content, /AI Technologies - Teams export/);
    assert.match(fullHistoryHtmlResult.content, /<h3>AI Discussion Topic<\/h3>/);

    await page.click(".tsm-export__launcher");
    await page.waitForFunction(() => document.body.classList.contains("tsm-export--selection-active"));

    const threadSelectors = await page.evaluate(
      () => Array.from(document.querySelectorAll('.tsm-export__message[data-tsm-message-id]')).length
    );
    assert.equal(threadSelectors, 3, "Should have 3 selectable thread elements");

    const firstMessageId = await page.evaluate(
      () => document.querySelector('.tsm-export__message[data-tsm-message-id]')?.dataset.tsmMessageId || ""
    );
    assert.ok(firstMessageId, "First thread should have a message ID");

    await page.click(`[data-tsm-message-id="${firstMessageId}"]`, { offset: { x: 40, y: 20 } });
    await page.waitForFunction(() => window.__teamsMessageExporter.state.selectedIds.size === 1);

    await page.click('.tsm-export__panel [data-action="copy-md"]');
    await page.waitForFunction(
      () => !document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open")
    );

    const copiedText = await page.evaluate(() => window.__copiedText || "");
    assert.match(copiedText, /Alice Example/, "Copied text should contain the thread author");
    assert.match(copiedText, /GPT-5 API/, "Copied text should contain the post body");
    assert.match(copiedText, /function calling is much improved/, "Copied text should contain the reply body");
  } finally {
    await browser.close();
  }
});
