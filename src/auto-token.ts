/**
 * Automated Teams Token Acquisition
 *
 * Launches a fresh Chrome instance, navigates to Teams, completes the
 * Microsoft Entra ID login flow using the macOS Intune/Company Portal
 * passkey provider, and captures the skype token via CDP Fetch interception.
 *
 * This enables fully automated token renewal with zero user interaction
 * on machines that have Intune/Company Portal configured as a platform
 * authenticator for FIDO2/WebAuthn.
 *
 * Requirements:
 *   - macOS with Intune/Company Portal configured as a passkey provider
 *   - System Chrome installed at /Applications/Google Chrome.app
 *   - playwright package installed
 *
 * Usage:
 *   import { acquireTokenAutomatically } from "./auto-token.js";
 *   const token = await acquireTokenAutomatically("maxim.mazurok@wisetechglobal.com");
 */

import { chromium, type BrowserContext, type CDPSession } from "playwright";
import type { TeamsAuthToken } from "./direct-api-client.js";

const TEAMS_URL = "https://teams.cloud.microsoft/";
const SYSTEM_CHROME_PATH = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const TEMP_PROFILE_DIRECTORY = "/tmp/teams-export-chrome-profile";
const LOGIN_TIMEOUT = 60_000;
const TOKEN_INTERCEPT_TIMEOUT = 30_000;

export interface AutoTokenOptions {
  /** Corporate email for Microsoft login */
  email: string;
  /** Path to Chrome executable (defaults to system Chrome on macOS) */
  chromePath?: string;
  /** Directory for temporary browser profile (cleaned on each run) */
  profileDirectory?: string;
  /** Run browser in headless mode (default: true) */
  headless?: boolean;
  /** Show progress messages (default: false) */
  verbose?: boolean;
}

/**
 * Acquire a Teams skype token fully automatically:
 *   1. Launch system Chrome with a fresh profile
 *   2. Navigate to Teams → redirects to Entra ID login
 *   3. Enter email → org redirects to FIDO/passkey
 *   4. macOS Intune completes passkey challenge silently
 *   5. Teams loads → capture skype token from network requests
 *   6. Close browser and return token
 */
export async function acquireTokenAutomatically(
  options: AutoTokenOptions,
): Promise<TeamsAuthToken> {
  const chromePath = options.chromePath ?? SYSTEM_CHROME_PATH;
  const profileDirectory = options.profileDirectory ?? TEMP_PROFILE_DIRECTORY;
  const headless = options.headless ?? true;
  const log = options.verbose ? console.log.bind(console) : () => {};

  // Clean up any previous profile to ensure a fresh session
  const { execSync } = await import("node:child_process");
  execSync(`rm -rf "${profileDirectory}"`);

  log("Launching Chrome with fresh profile...");
  const context = await chromium.launchPersistentContext(profileDirectory, {
    headless,
    executablePath: chromePath,
    args: [
      "--disable-blink-features=AutomationControlled",
      "--disable-features=VirtualAuthenticators",
      "--enable-features=WebAuthenticationMacPlatform",
      "--no-first-run",
      "--no-default-browser-check",
    ],
    ignoreDefaultArgs: [
      "--disable-component-extensions-with-background-pages",
    ],
  });

  try {
    const token = await performLoginAndCaptureToken(context, options.email, log);
    return token;
  } finally {
    await context.close();
    // Clean up profile directory
    try {
      execSync(`rm -rf "${profileDirectory}"`);
    } catch {
      // Best-effort cleanup
    }
  }
}

async function performLoginAndCaptureToken(
  context: BrowserContext,
  email: string,
  log: (...arguments_: unknown[]) => void,
): Promise<TeamsAuthToken> {
  const page = context.pages()[0] || await context.newPage();

  // Navigate to Teams → triggers redirect to Entra ID login
  log("Navigating to Teams...");
  await page.goto(TEAMS_URL, { waitUntil: "domcontentloaded", timeout: LOGIN_TIMEOUT });

  // Wait for the login page
  await page.waitForURL(/login\.microsoftonline\.com|login\.microsoft\.com/, {
    timeout: LOGIN_TIMEOUT,
  }).catch(() => {
    // May already be at Teams if session cached
  });

  // If we're at the login page, enter email
  if (page.url().includes("login.microsoftonline.com") || page.url().includes("login.microsoft.com")) {
    log("At login page, entering email...");

    const emailInput = await page.waitForSelector(
      '#i0116, input[name="loginfmt"], input[type="email"]',
      { timeout: 15_000 },
    ).catch(() => null);

    if (emailInput) {
      await emailInput.fill(email);

      const nextButton = await page.waitForSelector("#idSIButton9", { timeout: 5_000 }).catch(() => null);
      if (nextButton) {
        log("Submitting email...");
        await nextButton.click();
      }
    }

    // Wait for FIDO/passkey flow to complete and redirect back to Teams
    log("Waiting for passkey authentication...");
    await page.waitForURL(/teams\.cloud\.microsoft/, {
      timeout: LOGIN_TIMEOUT,
      waitUntil: "domcontentloaded",
    });
  }

  log(`Logged in. URL: ${page.url()}`);

  // Now capture the skype token by intercepting network requests via CDP
  log("Capturing skype token...");
  const token = await captureTokenViaCdp(page, log);

  return token;
}

async function captureTokenViaCdp(
  page: Awaited<ReturnType<BrowserContext["newPage"]>>,
  log: (...arguments_: unknown[]) => void,
): Promise<TeamsAuthToken> {
  const cdpSession: CDPSession = await page.context().newCDPSession(page);
  let skypeToken: string | null = null;

  await cdpSession.send("Fetch.enable", {
    patterns: [{ urlPattern: "*teams*", requestStage: "Request" }],
  });

  const tokenPromise = new Promise<void>((resolve) => {
    const timeout = setTimeout(resolve, TOKEN_INTERCEPT_TIMEOUT);

    cdpSession.on("Fetch.requestPaused", async (event: Record<string, unknown>) => {
      const request = event.request as { headers?: Record<string, string> };
      const requestId = event.requestId as string;

      for (const [name, value] of Object.entries(request.headers ?? {})) {
        if (name.toLowerCase() === "x-skypetoken" && !skypeToken) {
          skypeToken = value;
        }
      }

      try {
        await cdpSession.send("Fetch.continueRequest", { requestId });
      } catch {
        // Request may have already been handled
      }

      if (skypeToken) {
        clearTimeout(timeout);
        resolve();
      }
    });
  });

  // Reload the page to trigger authenticated API requests
  log("Reloading page to intercept token...");
  page.reload({ waitUntil: "domcontentloaded", timeout: LOGIN_TIMEOUT }).catch(() => {
    // Reload may time out, but token should be captured by then
  });

  await tokenPromise;

  await cdpSession.send("Fetch.disable");
  await cdpSession.detach();

  if (!skypeToken) {
    throw new Error(
      "Failed to capture skype token after login. Teams may not have fully loaded.",
    );
  }

  const capturedToken: string = skypeToken;
  log(`Token captured (${capturedToken.length} chars)`);

  return {
    skypeToken: capturedToken,
    region: "apac",
  };
}
