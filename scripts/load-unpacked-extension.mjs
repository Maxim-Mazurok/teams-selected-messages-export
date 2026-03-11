import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const extensionDir = path.join(rootDir, "extension-dist");
const browserUrl = `http://127.0.0.1:${Number(process.env.DEBUG_PORT || "9222")}`;
const extensionName = "Teams Selected Messages Export";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getExtensions(page) {
  return page.evaluate(async () => {
    const items = await chrome.management.getAll();
    return items.map((item) => ({
      name: item.name,
      id: item.id,
      enabled: item.enabled,
      installType: item.installType,
      type: item.type,
      hostPermissions: item.hostPermissions
    }));
  });
}

async function ensureExtensionsPage(browser) {
  const pages = await browser.pages();
  let page = pages.find((entry) => entry.url().startsWith("chrome://extensions"));

  if (!page) {
    page = await browser.newPage();
    await page.goto("chrome://extensions/", { waitUntil: "domcontentloaded" });
  }

  await page.waitForFunction(() => Boolean(document.querySelector("extensions-manager")), { timeout: 15000 });
  return page;
}

async function ensureDevMode(page) {
  await page.evaluate(() => {
    const manager = document.querySelector("extensions-manager");
    const toolbar = manager?.shadowRoot?.querySelector("extensions-toolbar");
    const devModeToggle = toolbar?.shadowRoot?.querySelector("#devMode");

    if (devModeToggle && !devModeToggle.checked) {
      devModeToggle.click();
    }
  });
}

async function clickLoadUnpacked(page) {
  const fileChooserPromise = page.waitForFileChooser({ timeout: 3000 }).catch(() => null);

  await page.evaluate(() => {
    const manager = document.querySelector("extensions-manager");
    const toolbar = manager?.shadowRoot?.querySelector("extensions-toolbar");
    const loadButton = toolbar?.shadowRoot?.querySelector("#loadUnpacked");
    loadButton?.click();
  });

  const fileChooser = await fileChooserPromise;
  if (fileChooser) {
    await fileChooser.accept([extensionDir]);
  }
}

async function waitForExtension(page, timeoutMs = 15000) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const items = await getExtensions(page);
    const match = items.find((item) => item.name === extensionName && item.installType === "development");

    if (match) {
      return { extension: match, items };
    }

    await sleep(500);
  }

  return { extension: null, items: await getExtensions(page) };
}

async function reloadExtension(page, extensionId) {
  await page.evaluate(
    (id) =>
      new Promise((resolve, reject) => {
        try {
          chrome.developerPrivate.reload(id, () => {
            const errorMessage = chrome.runtime.lastError?.message;
            if (errorMessage) {
              reject(new Error(errorMessage));
              return;
            }

            resolve();
          });
        } catch (error) {
          reject(error);
        }
      }),
    extensionId
  );
}

async function main() {
  const browser = await puppeteer.connect({ browserURL: browserUrl, defaultViewport: null });

  try {
    const page = await ensureExtensionsPage(browser);
    await ensureDevMode(page);

    const existing = await waitForExtension(page, 1000);
    if (existing.extension) {
      await reloadExtension(page, existing.extension.id);
      const refreshed = await waitForExtension(page, 5000);
      console.log(
        JSON.stringify({ browserUrl, extensionDir, extension: refreshed.extension || existing.extension, items: refreshed.items }, null, 2)
      );
      return;
    }

    await clickLoadUnpacked(page);
    const result = await waitForExtension(page);

    if (!result.extension) {
      throw new Error(
        `The unpacked extension did not appear in chrome://extensions on ${browserUrl}. ` +
          "Use npm run chrome:launch:extension if Chrome blocks the automated load flow."
      );
    }

    console.log(JSON.stringify({ browserUrl, extensionDir, extension: result.extension, items: result.items }, null, 2));
  } finally {
    await browser.disconnect();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
