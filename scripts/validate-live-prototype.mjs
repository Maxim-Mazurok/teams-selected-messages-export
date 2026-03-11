import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const sourcePath = path.join(rootDir, "src", "teams-export-prototype.js");
const artifactsDir = path.join(rootDir, "artifacts");
const exportsDir = path.join(artifactsDir, "exports");
const screenshotsDir = path.join(artifactsDir, "screenshots");
const debugPort = Number(process.env.DEBUG_PORT || "9222");
const browserUrl = `http://127.0.0.1:${debugPort}`;

function log(message) {
  process.stdout.write(`${message}\n`);
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function slug(value) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80);
}

function timestampPart() {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

async function ensureArtifactDirs() {
  await fs.mkdir(exportsDir, { recursive: true });
  await fs.mkdir(screenshotsDir, { recursive: true });
}

async function connectBrowser() {
  return puppeteer.connect({
    browserURL: browserUrl,
    defaultViewport: null
  });
}

async function getTeamsPage(browser) {
  const existingPages = await browser.pages();
  const teamsPage = existingPages.find((page) => /teams\.(microsoft|cloud\.microsoft)/i.test(page.url()));

  if (teamsPage) {
    await teamsPage.bringToFront();
    return teamsPage;
  }

  const page = await browser.newPage();
  await page.goto("https://teams.microsoft.com", { waitUntil: "domcontentloaded" });
  await page.bringToFront();
  return page;
}

async function waitForConversation(page) {
  log("Waiting for a Teams conversation to be visible...");
  await page.waitForFunction(
    () =>
      Boolean(
        document.querySelector('[data-tid="chat-pane-message"]') &&
          document.querySelector('[data-tid="chat-title"]')
      ),
    { timeout: 0 }
  );
}

async function injectPrototype(page) {
  const source = await fs.readFile(sourcePath, "utf8");

  await page.evaluate((script) => {
    eval(script);
  }, source);

  await page.waitForFunction(
    () =>
      Boolean(
        window.__teamsMessageExporter &&
          window.__teamsMessageExporter.state &&
          window.__teamsMessageExporter.state.messages.length > 0
      ),
    { timeout: 15000 }
  );

  return page.evaluate(() => ({
    conversationTitle: document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() || document.title,
    strategy: window.__teamsMessageExporter?.state?.strategy?.name || null,
    messageCount: window.__teamsMessageExporter?.state?.messages?.length || 0
  }));
}

async function listVisibleChats(page) {
  return page.evaluate(() => {
    const utilityLabels = new Set(["Mentions", "Followed threads", "Drafts"]);

    return Array.from(document.querySelectorAll('[role="treeitem"][aria-level="2"]'))
      .map((element, index) => {
        const text = (element.textContent || "").replace(/\s+/g, " ").trim();

        return {
          index,
          text
        };
      })
      .filter((chat) => chat.text && !utilityLabels.has(chat.text));
  });
}

async function openChat(page, chat) {
  const clicked = await page.evaluate((targetIndex) => {
    const candidates = Array.from(document.querySelectorAll('[role="treeitem"][aria-level="2"]'));
    const match = candidates[targetIndex];

    if (!match) {
      return false;
    }

    match.scrollIntoView({ block: "center" });
    match.click();
    return true;
  }, chat.index);

  if (!clicked) {
    return false;
  }

  await wait(2500);
  return true;
}

async function selectLatestMessages(page, count = 3) {
  const result = await page.evaluate((selectionCount) => {
    const exporter = window.__teamsMessageExporter;

    if (!exporter) {
      return { ok: false, reason: "Prototype API not found." };
    }

    exporter.setActive(true);
    exporter.clearSelection();

    const sample = exporter.state.messages.slice(-selectionCount);
    if (sample.length < 2) {
      return { ok: false, reason: "Not enough visible messages to validate range selection." };
    }

    const clickMessage = (messageId, shiftKey = false) => {
      const selector = `[data-tsm-message-id="${CSS.escape(messageId)}"]`;
      const element = document.querySelector(selector);
      if (!element) {
        throw new Error(`Unable to find message element for ${messageId}`);
      }

      element.dispatchEvent(
        new MouseEvent("click", {
          bubbles: true,
          cancelable: true,
          shiftKey,
          view: window
        })
      );
    };

    clickMessage(sample[0].id);
    clickMessage(sample[sample.length - 1].id, true);

    const selectedMessages = exporter.getSelectedMessages();

    return {
      ok: true,
      selectedCount: selectedMessages.length,
      selectedIds: selectedMessages.map((message) => message.id),
      sampleIds: sample.map((message) => message.id),
      conversationTitle: document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() || document.title,
      markdown: exporter.renderMarkdown(),
      html: exporter.renderHtml(),
      messageCount: exporter.state.messages.length
    };
  }, count);

  return result;
}

async function saveValidationArtifacts(result, kind) {
  const baseName = `${slug(result.conversationTitle || "teams-conversation")}-${kind}-${timestampPart()}`;
  const markdownPath = path.join(exportsDir, `${baseName}.md`);
  const htmlPath = path.join(exportsDir, `${baseName}.html`);

  await fs.writeFile(markdownPath, result.markdown, "utf8");
  await fs.writeFile(htmlPath, result.html, "utf8");

  return {
    markdownPath,
    htmlPath
  };
}

async function screenshot(page, conversationTitle, kind) {
  const baseName = `${slug(conversationTitle || "teams-conversation")}-${kind}-${timestampPart()}.png`;
  const targetPath = path.join(screenshotsDir, baseName);
  await page.screenshot({ path: targetPath, fullPage: false });
  return targetPath;
}

async function validateConversation(page, kind) {
  const injected = await injectPrototype(page);
  log(`Injected prototype into "${injected.conversationTitle}" (${injected.messageCount} visible messages).`);

  const selection = await selectLatestMessages(page, 3);
  if (!selection.ok) {
    return {
      ok: false,
      kind,
      conversationTitle: injected.conversationTitle,
      reason: selection.reason
    };
  }

  const files = await saveValidationArtifacts(selection, kind);
  const screenshotPath = await screenshot(page, selection.conversationTitle, kind);

  return {
    ok: true,
    kind,
    conversationTitle: selection.conversationTitle,
    visibleMessageCount: selection.messageCount,
    selectedCount: selection.selectedCount,
    files,
    screenshotPath
  };
}

async function main() {
  await ensureArtifactDirs();

  const browser = await connectBrowser();

  try {
    const page = await getTeamsPage(browser);
    await waitForConversation(page);

    const validations = [];
    validations.push(await validateConversation(page, "current"));

    const currentTitle = validations[0].conversationTitle;
    const chats = await listVisibleChats(page);
    const alternateChat = chats.find((chat) => !chat.text.includes(currentTitle || ""));

    if (alternateChat) {
      log(`Switching to alternate chat: ${alternateChat.text}`);
      const opened = await openChat(page, alternateChat);

      if (opened) {
        await waitForConversation(page);
        validations.push(await validateConversation(page, "alternate"));
      } else {
        validations.push({
          ok: false,
          kind: "alternate",
          conversationTitle: alternateChat.text,
          reason: "Unable to click alternate chat in the sidebar."
        });
      }
    } else {
      validations.push({
        ok: false,
        kind: "alternate",
        conversationTitle: null,
        reason: "No alternate visible chat candidate found in the current sidebar."
      });
    }

    const summary = {
      generatedAt: new Date().toISOString(),
      browserUrl,
      validations
    };

    const summaryPath = path.join(artifactsDir, "validation-summary.json");
    await fs.writeFile(summaryPath, JSON.stringify(summary, null, 2), "utf8");

    log(`Validation summary written to ${summaryPath}`);
    log(JSON.stringify(summary, null, 2));
  } finally {
    await browser.disconnect();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
