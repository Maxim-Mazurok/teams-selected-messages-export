import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import puppeteer from "puppeteer-core";
import { findChromeExecutable } from "../helpers/chrome-path.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDirectory = path.resolve(__dirname, "../..");
const sourcePath = path.join(rootDirectory, "dist", "content-script.js");
const fixturePath = path.join(
  rootDirectory,
  "tests",
  "fixtures",
  "teams-thread-replies-fixture.html"
);

async function injectPrototype(page) {
  const source = await fs.readFile(sourcePath, "utf8");
  await page.evaluate((script) => {
    eval(script);
  }, source);
  await page.waitForFunction(
    () => window.__teamsMessageExporter?.state?.messages?.length >= 4,
    {
      timeout: 10000
    }
  );
}

test("thread replies view supports selection and full-history export includes original post", async () => {
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

    await page.goto(pathToFileURL(fixturePath).href, { waitUntil: "load" });
    await injectPrototype(page);

    const initialState = await page.evaluate(() => ({
      messageCount: window.__teamsMessageExporter.state.messages.length,
      strategyName: window.__teamsMessageExporter.state.strategy?.name || ""
    }));

    assert.equal(initialState.messageCount, 4);
    assert.equal(initialState.strategyName, "teams-channel-pane-message");

    await page.click(".tsm-export__launcher");
    await page.waitForFunction(() => document.body.classList.contains("tsm-export--selection-active"));

    await page.click('[data-tsm-message-id="r3"]', { offset: { x: 40, y: 20 } });
    await page.waitForFunction(() => window.__teamsMessageExporter.state.selectedIds.size === 1);

    await page.click('.tsm-export__panel [data-action="copy-md"]');
    await page.waitForFunction(
      () => !document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open")
    );

    const copiedText = await page.evaluate(() => window.__copiedText || "");
    assert.match(copiedText, /Mina Chen/);
    assert.match(copiedText, /flaky integration tests/);

    const fullHistoryMarkdown = await page.evaluate(async () => {
      const result = await window.__teamsMessageExporter.exportFullHistory("md", {
        download: false,
        closePanel: false
      });

      return {
        count: result?.count ?? 0,
        content: result?.content ?? ""
      };
    });

    assert.equal(fullHistoryMarkdown.count, 4);
    assert.match(fullHistoryMarkdown.content, /Full chat history messages: 4/);
    assert.match(fullHistoryMarkdown.content, /## Priya Patel \| 09:18/);
    assert.match(fullHistoryMarkdown.content, /standardize our AI coding review checklist/);
    assert.match(fullHistoryMarkdown.content, /## Alex Jordan \| 09:19/);
    assert.match(fullHistoryMarkdown.content, /## Mina Chen \| 09:21/);
    assert.match(fullHistoryMarkdown.content, /## Leo Singh \| 09:24/);

    const fullHistoryHtml = await page.evaluate(async () => {
      const result = await window.__teamsMessageExporter.exportFullHistory("html", {
        download: false,
        closePanel: false
      });

      return {
        count: result?.count ?? 0,
        content: result?.content ?? "",
        format: result?.format ?? ""
      };
    });

    assert.equal(fullHistoryHtml.count, 4);
    assert.equal(fullHistoryHtml.format, "html");
    assert.match(fullHistoryHtml.content, /<!doctype html>/i);
    assert.match(fullHistoryHtml.content, /AI for Developers - Teams export/);
    assert.match(fullHistoryHtml.content, /Priya Patel/);
    assert.match(fullHistoryHtml.content, /standardize our AI coding review checklist/);
  } finally {
    await browser.close();
  }
});
