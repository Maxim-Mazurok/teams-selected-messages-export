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
const sourcePath = path.join(rootDir, "src", "teams-export-prototype.js");
const fixturePath = path.join(rootDir, "tests", "fixtures", "teams-fixture.html");

async function injectPrototype(page) {
  const source = await fs.readFile(sourcePath, "utf8");
  await page.evaluate((script) => {
    eval(script);
  }, source);
  await page.waitForFunction(() => window.__teamsMessageExporter?.state?.messages?.length >= 3, {
    timeout: 10000
  });
}

test("fixture flow keeps panel copy contextual, keeps the action grid stable, and closes after panel copy", async () => {
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
    await page.waitForFunction(() => document.querySelector(".tsm-export__dock")?.dataset.theme === "dark");

    assert.equal(
      await page.$eval(".tsm-export__dock", (element) => element.parentElement?.id || ""),
      "titlebar-end-slot-start"
    );
    assert.equal(
      await page.$eval(".tsm-export__dock", (element) => element.dataset.theme),
      "dark"
    );

    await page.evaluate(() => {
      document.documentElement.className = "theme-defaultV2";
    });
    await page.waitForFunction(() => document.querySelector(".tsm-export__dock")?.dataset.theme === "light");

    assert.equal(
      await page.$eval(".tsm-export__dock", (element) => element.dataset.theme),
      "light"
    );

    await page.click(".tsm-export__launcher");
    await page.waitForFunction(() => document.body.classList.contains("tsm-export--selection-active"));

    await page.click('[data-tsm-message-id="m1"]', { offset: { x: 40, y: 20 } });
    await page.keyboard.down("Shift");
    await page.click('[data-tsm-message-id="m3"]', { offset: { x: 40, y: 20 } });
    await page.keyboard.up("Shift");

    await page.waitForFunction(() => document.querySelectorAll(".tsm-export__message--selected").length === 3);

    const selectedMessageVisuals = await page.$eval('[data-tsm-message-id="m2"]', (element) => {
      const styles = getComputedStyle(element);
      return {
        backgroundColor: styles.backgroundColor
      };
    });

    assert.notEqual(selectedMessageVisuals.backgroundColor, "rgba(0, 0, 0, 0)");

    assert.equal(
      await page.$eval('.tsm-export__panel [data-action="copy-md"]', (element) => element.hidden),
      false
    );
    assert.equal(
      await page.$eval('.tsm-export__panel [data-action="copy-md"]', (element) => element.disabled),
      false
    );
    assert.equal(
      await page.$eval('.tsm-export__panel [data-role="description"]', (element) => getComputedStyle(element).whiteSpace),
      "normal"
    );

    await page.click('.tsm-export__panel [data-role="active-toggle"]');
    await page.waitForFunction(() => !document.body.classList.contains("tsm-export--selection-active"));

    const hiddenState = await page.evaluate(() => ({
      panelCopyHidden: document.querySelector('.tsm-export__panel [data-action="copy-md"]').hidden,
      panelCopyDisabled: document.querySelector('.tsm-export__panel [data-action="copy-md"]').disabled,
      actionOrder: Array.from(document.querySelectorAll('.tsm-export__panel [data-role="selection-actions"] > button')).map(
        (element) => element.dataset.action
      ),
      actionLabels: Array.from(document.querySelectorAll('.tsm-export__panel [data-role="selection-actions"] > button')).map(
        (element) => element.textContent?.trim() || ""
      ),
      selectedVisibleCount: document.querySelectorAll(".tsm-export__message--selected").length
    }));

    assert.equal(hiddenState.panelCopyHidden, false);
    assert.equal(hiddenState.panelCopyDisabled, true);
    assert.deepEqual(hiddenState.actionOrder, [
      "copy-md",
      "clear",
      "export-md",
      "export-html",
      "export-full-md",
      "export-full-html"
    ]);
    assert.deepEqual(hiddenState.actionLabels, [
      "Copy MD",
      "Clear Selection",
      "Download MD",
      "Download HTML",
      "Full chat MD",
      "Full chat HTML"
    ]);
    assert.equal(hiddenState.selectedVisibleCount, 0);

    await page.click('.tsm-export__panel [data-role="active-toggle"]');
    await page.waitForFunction(
      () =>
        document.body.classList.contains("tsm-export--selection-active") &&
        document.querySelectorAll(".tsm-export__message--selected").length === 3
    );
    assert.equal(
      await page.$eval('.tsm-export__panel [data-action="copy-md"]', (element) => element.disabled),
      false
    );

    await page.click('.tsm-export__panel [data-action="copy-md"]');
    await page.waitForFunction(
      () => !document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open")
    );

    const copiedState = await page.evaluate(() => ({
      copiedText: window.__copiedText || "",
      panelOpen: document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open") || false,
      active: document.body.classList.contains("tsm-export--selection-active")
    }));

    assert.equal(copiedState.panelOpen, false);
    assert.equal(copiedState.active, false);
    assert.match(copiedState.copiedText, /@Maxim Mazurok/);
    assert.match(copiedState.copiedText, /> Replying to Alice Example \| 08:58/);
    assert.match(copiedState.copiedText, /\[Image omitted\]/);
    assert.match(copiedState.copiedText, /Reactions: 👍 like by Jonathan Clegg/);

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

    assert.equal(fullHistoryResult.count, 3);
    assert.match(fullHistoryResult.content, /Full chat history messages: 3/);
    assert.match(fullHistoryResult.content, /## Alice Example \| 08:58/);
    assert.match(fullHistoryResult.content, /## Bob Example \| 09:00/);
    assert.match(fullHistoryResult.content, /## Maxim Mazurok \| 09:02/);

    const fullHistoryHtmlResult = await page.evaluate(async () => {
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

    assert.equal(fullHistoryHtmlResult.count, 3);
    assert.equal(fullHistoryHtmlResult.format, "html");
    assert.match(fullHistoryHtmlResult.content, /<!doctype html>/i);
    assert.match(fullHistoryHtmlResult.content, /Fixture Chat - Teams export/);

    await page.evaluate(() => {
      document.querySelector('[data-tid="chat-title"]').textContent = "Second Fixture Chat";
      document.querySelector('[data-tid="message-pane-list-runway"]').innerHTML = `
        <div data-tid="chat-pane-item" data-message-id="m4">
          <div data-testid="message-wrapper">
            <div data-tid="message-author-name">Charlie Example</div>
            <time datetime="2026-03-11T09:10:00.000Z">09:10</time>
            <div data-tid="chat-pane-message">
              <p>Fresh chat, first message.</p>
            </div>
          </div>
        </div>
        <div data-tid="chat-pane-item" data-message-id="m5">
          <div data-testid="message-wrapper">
            <div data-tid="message-author-name">Dana Example</div>
            <time datetime="2026-03-11T09:11:00.000Z">09:11</time>
            <div data-tid="chat-pane-message">
              <p>Fresh chat, second message.</p>
            </div>
          </div>
        </div>
      `;
    });
    await page.evaluate(() => {
      window.__teamsMessageExporter.refresh();
    });

    await page.waitForFunction(
      () => window.__teamsMessageExporter.state.conversationKey.includes("Second Fixture Chat"),
      { timeout: 30000 }
    );
    await page.waitForFunction(() => window.__teamsMessageExporter.state.selectedIds.size === 0, {
      timeout: 30000
    });
    await page.waitForFunction(() => document.querySelectorAll('.tsm-export__message[data-tsm-message-id]').length === 2, {
      timeout: 30000
    });

    const switchedChatState = await page.evaluate(() => ({
      title: document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() || "",
      selectedIds: window.__teamsMessageExporter.state.selectedIds.size,
      visibleSelectedCount: document.querySelectorAll(".tsm-export__message--selected").length,
      countLabel: document.querySelector('.tsm-export__panel [data-role="count"]')?.textContent?.trim() || null
    }));

    assert.equal(switchedChatState.title, "Second Fixture Chat");
    assert.equal(switchedChatState.selectedIds, 0);
    assert.equal(switchedChatState.visibleSelectedCount, 0);
    assert.equal(switchedChatState.countLabel, "0 selected");

    await page.click(".tsm-export__launcher");
    await page.waitForFunction(() => document.body.classList.contains("tsm-export--selection-active"));
    await page.click('[data-tsm-message-id="m4"]', { offset: { x: 40, y: 20 } });
    await page.keyboard.down("Shift");
    await page.click('[data-tsm-message-id="m5"]', { offset: { x: 40, y: 20 } });
    await page.keyboard.up("Shift");
    await page.waitForFunction(() => document.querySelectorAll(".tsm-export__message--selected").length === 2);

    await page.click('.tsm-export__panel [data-action="copy-md"]');
    await page.waitForFunction(
      () => !document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open")
    );

    const switchedCopy = await page.evaluate(() => window.__copiedText || "");
    assert.match(switchedCopy, /# Second Fixture Chat/);
    assert.match(switchedCopy, /Charlie Example/);
    assert.match(switchedCopy, /Dana Example/);
  } finally {
    await browser.close();
  }
});
