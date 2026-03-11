import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const artifactsDir = path.join(rootDir, "artifacts");
const exportsDir = path.join(artifactsDir, "exports");
const screenshotsDir = path.join(artifactsDir, "screenshots");
const summaryPath = path.join(artifactsDir, "installed-extension-validation.json");
const browserUrl = `http://127.0.0.1:${Number(process.env.DEBUG_PORT || "9222")}`;

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

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function ensureDirs() {
  await fs.mkdir(exportsDir, { recursive: true });
  await fs.mkdir(screenshotsDir, { recursive: true });
}

async function clickCenter(page, selector) {
  const rect = await page.$eval(selector, (element) => {
    const box = element.getBoundingClientRect();
    return {
      x: box.x + box.width / 2,
      y: box.y + box.height / 2
    };
  });

  await page.mouse.click(rect.x, rect.y);
}

async function waitForDownloads(downloadDir, extension, timeoutMs = 15000) {
  return waitForNewDownload(downloadDir, extension, new Set(), timeoutMs);
}

async function waitForNewDownload(downloadDir, extension, knownFiles = new Set(), timeoutMs = 15000) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const entries = await fs.readdir(downloadDir, { withFileTypes: true });
    const completeFiles = entries
      .filter((entry) => entry.isFile() && entry.name.endsWith(extension) && !entry.name.endsWith(".crdownload"))
      .map((entry) => path.join(downloadDir, entry.name));

    const newFiles = completeFiles.filter((filePath) => !knownFiles.has(filePath));
    if (newFiles.length > 0) {
      return newFiles.sort().at(-1);
    }

    await sleep(250);
  }

  throw new Error(`Timed out waiting for a ${extension} download in ${downloadDir}`);
}

async function listCompleteDownloads(downloadDir, extension) {
  const entries = await fs.readdir(downloadDir, { withFileTypes: true });
  return new Set(
    entries
      .filter((entry) => entry.isFile() && entry.name.endsWith(extension) && !entry.name.endsWith(".crdownload"))
      .map((entry) => path.join(downloadDir, entry.name))
  );
}

async function enableDownloads(page, downloadDir) {
  const client = await page.target().createCDPSession();

  try {
    await client.send("Browser.setDownloadBehavior", {
      behavior: "allow",
      downloadPath: downloadDir
    });
  } catch {
    await client.send("Page.setDownloadBehavior", {
      behavior: "allow",
      downloadPath: downloadDir
    });
  }
}

async function ensurePanelOpen(page) {
  const panelOpen = await page.evaluate(
    () => document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open") || false
  );

  if (!panelOpen) {
    await clickCenter(page, ".tsm-export__launcher");
    await page.waitForFunction(
      () => document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open") || false,
      { timeout: 5000 }
    );
  }
}

async function selectLatestMessages(page, count = 3) {
  await clickCenter(page, '.tsm-export__panel [data-action="clear"]');
  await sleep(200);

  const targets = await page.evaluate((selectionCount) => {
    return Array.from(document.querySelectorAll('.tsm-export__message[data-tsm-message-id]'))
      .slice(-selectionCount)
      .map((element) => {
        const rect = element.getBoundingClientRect();
        return {
          id: element.dataset.tsmMessageId,
          point: {
            x: rect.x + 24,
            y: rect.y + Math.min(24, Math.max(rect.height / 2, 12))
          },
          text: (element.textContent || "").replace(/\s+/g, " ").trim().slice(0, 120)
        };
      })
      .filter((target) => Number.isFinite(target.point.x) && Number.isFinite(target.point.y));
  }, count);

  if (targets.length < 2) {
    throw new Error(`Not enough decorated messages were available for validation (${targets.length}).`);
  }

  await page.mouse.click(targets[0].point.x, targets[0].point.y);
  await sleep(150);
  await page.keyboard.down("Shift");
  await page.mouse.click(targets[targets.length - 1].point.x, targets[targets.length - 1].point.y);
  await page.keyboard.up("Shift");

  await page.waitForFunction(
    (expectedCount) => document.querySelectorAll(".tsm-export__message--selected").length >= expectedCount,
    { timeout: 5000 },
    targets.length
  );

  return page.evaluate(() => ({
    conversationTitle: document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() || document.title,
    selectedCount: document.querySelectorAll(".tsm-export__message--selected").length,
    countLabel: document.querySelector('.tsm-export__panel [data-role="count"]')?.textContent?.trim() || null,
    quickCopyVisible: !(document.querySelector(".tsm-export__quick-copy")?.hidden ?? true)
  }));
}

async function main() {
  await ensureDirs();
  const browser = await puppeteer.connect({ browserURL: browserUrl, defaultViewport: null });

  try {
    const pages = await browser.pages();
    const page = pages.find((entry) => /teams\.(microsoft|cloud\.microsoft)/i.test(entry.url()));

    if (!page) {
      throw new Error("No Teams page found in the connected Chrome instance.");
    }

    await browser.defaultBrowserContext().overridePermissions(new URL(page.url()).origin, [
      "clipboard-read",
      "clipboard-write"
    ]);

    await page.bringToFront();
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.waitForFunction(
      () =>
        Boolean(
          document.querySelector('[data-tid="chat-title"]') &&
            document.querySelector(".tsm-export__launcher") &&
            document.querySelectorAll('.tsm-export__message[data-tsm-message-id]').length >= 3
        ),
      { timeout: 30000 }
    );

    await ensurePanelOpen(page);

    const autoActivatedOnPanelOpen = await page.evaluate(() =>
      document.body.classList.contains("tsm-export--selection-active")
    );

    if (!autoActivatedOnPanelOpen) {
      throw new Error("Opening the export panel did not automatically enable selection mode.");
    }

    const selection = await selectLatestMessages(page, 3);

    await clickCenter(page, ".tsm-export__launcher");
    await page.waitForFunction(() => !(document.querySelector(".tsm-export__quick-copy")?.hidden ?? true), {
      timeout: 5000
    });
    const quickCopyState = await page.evaluate(() => ({
      visible: !(document.querySelector(".tsm-export__quick-copy")?.hidden ?? true),
      label: document.querySelector(".tsm-export__quick-copy")?.textContent?.trim() || null
    }));
    await clickCenter(page, ".tsm-export__quick-copy");
    await sleep(500);

    const clipboard = await page.evaluate(() => navigator.clipboard.readText());
    const postCopyState = await page.evaluate(() => ({
      active: document.body.classList.contains("tsm-export--selection-active"),
      selectedVisibleCount: document.querySelectorAll(".tsm-export__message--selected").length,
      quickCopyVisible: !(document.querySelector(".tsm-export__quick-copy")?.hidden ?? true)
    }));

    await clickCenter(page, ".tsm-export__launcher");
    await page.waitForFunction(() => {
      const panelOpen = document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open") || false;
      const active = document.body.classList.contains("tsm-export--selection-active");
      const restored = document.querySelectorAll(".tsm-export__message--selected").length;
      return panelOpen && active && restored >= 3;
    }, { timeout: 5000 });
    const restoredState = await page.evaluate(() => ({
      active: document.body.classList.contains("tsm-export--selection-active"),
      restoredVisibleCount: document.querySelectorAll(".tsm-export__message--selected").length,
      countLabel: document.querySelector('.tsm-export__panel [data-role="count"]')?.textContent?.trim() || null
    }));

    const baseName = `${slug(selection.conversationTitle)}-installed-extension-${timestampPart()}`;
    const downloadDir = path.join(exportsDir, `${baseName}-downloads`);
    const screenshotPath = path.join(screenshotsDir, `${baseName}.png`);
    await fs.mkdir(downloadDir, { recursive: true });
    await enableDownloads(page, downloadDir);

    await clickCenter(page, '.tsm-export__panel [data-action="export-md"]');
    const markdownPath = await waitForDownloads(downloadDir, ".md");

    await clickCenter(page, '.tsm-export__panel [data-action="export-html"]');
    const htmlPath = await waitForDownloads(downloadDir, ".html");

    await page.screenshot({ path: screenshotPath, fullPage: false });

    await clickCenter(page, '.tsm-export__panel [data-action="copy-md"]');
    await page.waitForFunction(
      () => !(document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open") || false),
      { timeout: 5000 }
    );
    const panelCopyState = await page.evaluate(() => ({
      panelOpen: document.querySelector(".tsm-export__panel")?.classList.contains("tsm-export__panel--open") || false,
      active: document.body.classList.contains("tsm-export--selection-active")
    }));

    await ensurePanelOpen(page);
    const existingMarkdownFiles = await listCompleteDownloads(downloadDir, ".md");
    await clickCenter(page, '.tsm-export__panel [data-action="export-full-md"]');
    const fullChatMarkdownPath = await waitForNewDownload(downloadDir, ".md", existingMarkdownFiles, 45000);
    const fullChatMarkdown = await fs.readFile(fullChatMarkdownPath, "utf8");
    const fullChatHeadingCount = (fullChatMarkdown.match(/^## /gm) || []).length;
    const fullChatScopePresent = /Full chat history messages:/m.test(fullChatMarkdown);

    const result = {
      browserUrl,
      conversationTitle: selection.conversationTitle,
      selectedCount: selection.selectedCount,
      countLabel: selection.countLabel,
      quickCopyVisible: quickCopyState.visible,
      quickCopyLabel: quickCopyState.label,
      autoActivatedOnPanelOpen,
      selectionStoppedAfterCopy: !postCopyState.active,
      visibleSelectionUiAfterCopy: postCopyState.selectedVisibleCount,
      selectionRestoredAfterReopen: restoredState.active,
      restoredVisibleCount: restoredState.restoredVisibleCount,
      restoredCountLabel: restoredState.countLabel,
      panelClosedAfterPanelCopy: !panelCopyState.panelOpen,
      selectionStoppedAfterPanelCopy: !panelCopyState.active,
      fullChatMarkdownPath,
      fullChatHeadingCount,
      fullChatScopePresent,
      clipboardPreview: clipboard.slice(0, 500),
      markdownPath,
      htmlPath,
      screenshotPath
    };

    await fs.writeFile(summaryPath, JSON.stringify(result, null, 2), "utf8");
    console.log(JSON.stringify({ ...result, summaryPath }, null, 2));
  } finally {
    await browser.disconnect();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
