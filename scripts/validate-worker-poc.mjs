/**
 * Full POC Validation — Worker Intercept
 *
 * Validates the complete worker bridge pipeline end-to-end:
 *   1. document_start MAIN world hook patches window.Worker
 *   2. Worker messages are intercepted and summarised
 *   3. Bridge (window.postMessage) sends data to isolated content script
 *   4. Worker store in isolated world receives and stores messages
 *   5. Member identity map is populated from message authors
 *
 * Requires: Chrome on port 9223 with extension loaded and Teams authenticated.
 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const artifactsDir = path.join(rootDir, "artifacts");

const BROWSER_URL = "http://127.0.0.1:9223";
const EXTENSION_NAME = "Teams Selected Messages Export";

function sleep(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function reloadExtension(extensionsPage, extensionName) {
  const extensionId = await extensionsPage.evaluate((name) =>
    new Promise((resolve, reject) => {
      chrome.management.getAll((items) => {
        const match = items.find((item) => item.name === name);
        if (match) resolve(match.id);
        else reject(new Error(`Extension "${name}" not found`));
      });
    })
  , extensionName);

  await extensionsPage.evaluate((identifier) =>
    new Promise((resolve, reject) => {
      chrome.developerPrivate.reload(identifier, () => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve();
        }
      });
    })
  , extensionId);

  return extensionId;
}

async function main() {
  await fs.mkdir(artifactsDir, { recursive: true });

  const browser = await puppeteer.connect({ browserURL: BROWSER_URL, defaultViewport: null });

  try {
    let pages = await browser.pages();
    let extensionsPage = pages.find((page) => page.url().startsWith("chrome://extensions"));
    if (!extensionsPage) {
      extensionsPage = await browser.newPage();
      await extensionsPage.goto("chrome://extensions/", { waitUntil: "domcontentloaded" });
      await sleep(2000);
    }

    console.log("Reloading extension with latest build...");
    const extensionId = await reloadExtension(extensionsPage, EXTENSION_NAME);
    console.log(`  Extension ID: ${extensionId}`);
    await sleep(1500);

    pages = await browser.pages();
    const teamsPage = pages.find((page) => /teams/.test(page.url()));
    if (!teamsPage) throw new Error("Teams page not found");
    await teamsPage.bringToFront();

    const consoleLogsFromExtension = [];
    teamsPage.on("console", (message) => {
      const text = message.text();
      if (text.includes("[teams-export]") || text.includes("[worker")) {
        consoleLogsFromExtension.push(`[${message.type()}] ${text}`);
      }
    });

    console.log("Reloading Teams page to trigger document_start hook...");
    await teamsPage.reload({ waitUntil: "domcontentloaded" });

    // Install bridge capture listener immediately after reload
    await teamsPage.evaluate(() => {
      window.__pocBridgeCapture = [];
      window.addEventListener("message", (event) => {
        if (event.source !== window || event.data?.source !== "teams-export-worker-hook") return;
        const entry = {
          type: event.data.type,
          ts: new Date().toISOString(),
          operationName: event.data.payload?.operationName || null,
          dataSource: event.data.payload?.dataSource || null,
          messageCount: Array.isArray(event.data.payload?.messages)
            ? event.data.payload.messages.length
            : null,
          memberCount: Array.isArray(event.data.payload?.members)
            ? event.data.payload.members.length
            : null,
          meId: event.data.payload?.id || null,
        };
        // Capture sample messages from first batch only
        if (
          !window.__pocBridgeCapture.some((item) => item.sampleMessages) &&
          Array.isArray(event.data.payload?.messages)
        ) {
          entry.sampleMessages = event.data.payload.messages.slice(0, 3).map((message) => ({
            id: message.id,
            author: message.author,
            authorId: message.authorId ? `...${String(message.authorId).slice(-8)}` : null,
            type: message.messageType,
            hasEmotions: Array.isArray(message.emotions) && message.emotions.length > 0,
            emotionsSummary: message.emotionsSummary,
            emotions: message.emotions,
          }));
        }
        window.__pocBridgeCapture.push(entry);
      });
    });

    console.log("Waiting 14 seconds for Teams workers to post messages...");
    await sleep(14000);

    const bridgeCapture = await teamsPage.evaluate(() => window.__pocBridgeCapture || []);
    const diagnostics = await teamsPage.evaluate(() => ({
      url: location.href,
      title: document.title,
      chatTitle: document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() || null,
      workerHookInstalled: Boolean(window.__teamsExportWorkerHookInstalled),
      workerName: window.Worker?.name,
      launcherExists: Boolean(document.querySelector(".tsm-export__launcher")),
      chatPaneItems: document.querySelectorAll('[data-tid="chat-pane-item"]').length,
      workerMessagesAttr: document.documentElement.dataset.teamsExportWorkerMessages || null,
      workerMembersAttr: document.documentElement.dataset.teamsExportWorkerMembers || null,
    }));

    console.log("\nBridge events from MAIN world:");
    bridgeCapture.forEach((event) => {
      console.log(
        `   [${event.type}] op=${event.operationName || "?"} ` +
          `msgs=${event.messageCount != null ? event.messageCount : "?"} ` +
          `source=${event.dataSource || "?"}`
      );
    });

    const sampleBatch = bridgeCapture.find((event) => event.sampleMessages);
    if (sampleBatch) {
      console.log("\nSample messages from worker batch:");
      sampleBatch.sampleMessages.forEach((message) => {
        console.log(
          `   id=${message.id} author="${message.author}" ` +
            `reactions=${JSON.stringify(message.emotionsSummary || [])} ` +
            `hasEmotions=${message.hasEmotions}`
        );
      });
    }

    console.log("\nPage diagnostics:", JSON.stringify(diagnostics, null, 2));
    console.log("\nExtension console logs:", consoleLogsFromExtension.slice(0, 10));

    const totalBridgeMessages = bridgeCapture.reduce(
      (total, event) => total + (event.messageCount || 0),
      0
    );

    const assessment = {
      workerHookInstalled: diagnostics.workerHookInstalled,
      workerConstructorPatched: diagnostics.workerName === "WorkerWrapper",
      bridgeWorking: bridgeCapture.length > 0,
      totalBridgeEvents: bridgeCapture.length,
      totalMessagesFromWorker: totalBridgeMessages,
      isolatedWorldReceivingMessages: Boolean(diagnostics.workerMessagesAttr),
      isolatedWorldMessageCount: diagnostics.workerMessagesAttr
        ? Number(diagnostics.workerMessagesAttr)
        : 0,
      membersIdentifiedInIsolatedWorld: diagnostics.workerMembersAttr
        ? Number(diagnostics.workerMembersAttr)
        : 0,
      extensionUiAvailable: diagnostics.launcherExists,
    };

    const result = {
      timestamp: new Date().toISOString(),
      extensionId,
      chatTitle: diagnostics.chatTitle,
      bridgeCapture: bridgeCapture.map(({ sampleMessages: _sampleMessages, ...rest }) => rest),
      bridgeSampleMessages: sampleBatch?.sampleMessages || [],
      diagnostics,
      consoleLogsFromExtension,
      assessment,
    };

    const resultPath = path.join(artifactsDir, "worker-poc-validation.json");
    await fs.writeFile(resultPath, JSON.stringify(result, null, 2), "utf8");

    console.log("\n=== POC VALIDATION ASSESSMENT ===");
    Object.entries(assessment).forEach(([key, value]) => {
      const symbol = value === false || value === 0 ? "FAIL" : "PASS";
      console.log(`  [${symbol}] ${key}: ${value}`);
    });
    console.log(`\nFull results: ${resultPath}`);
  } finally {
    await browser.disconnect();
  }
}

main().catch((error) => {
  console.error("Validation failed:", error.message);
  process.exitCode = 1;
});
