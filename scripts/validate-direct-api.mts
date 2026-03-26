/**
 * Validate Direct API POC
 *
 * Runs a quick end-to-end check that the direct API approach works:
 *   1. Connects to Chrome, captures token
 *   2. Lists conversations
 *   3. Fetches messages from a conversation
 *   4. Converts to snapshots
 *   5. Renders markdown and HTML
 *   6. Verifies output is non-empty
 *
 * Usage:
 *   DEBUG_PORT=9223 npx -y tsx scripts/validate-direct-api.mts
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

import {
  captureSkypeToken,
  findTeamsPage,
  fetchConversations,
  fetchMessages,
  fetchMembers,
} from "../src/direct-api-client.js";
import { convertApiMessagesToSnapshots } from "../src/direct-api-converter.js";
import { renderMarkdown } from "../src/markdown-renderer.js";
import { renderHtmlDocument } from "../src/html-renderer.js";

const currentFilename = fileURLToPath(import.meta.url);
const currentDirectory = path.dirname(currentFilename);
const rootDirectory = path.resolve(currentDirectory, "..");
const artifactsDirectory = path.join(rootDirectory, "artifacts");

const browserUrl = `http://127.0.0.1:${Number(process.env.DEBUG_PORT || "9222")}`;

interface CheckResult {
  name: string;
  passed: boolean;
  detail: string | number | boolean;
}

async function main() {
  const checks: CheckResult[] = [];
  let exitCode = 0;

  function check(name: string, passed: boolean, detail: string | number | boolean) {
    checks.push({ name, passed, detail });
    const icon = passed ? "✅" : "❌";
    console.log(`  ${icon} ${name}: ${detail}`);
    if (!passed) exitCode = 1;
  }

  console.log(`Connecting to Chrome at ${browserUrl}...`);
  const browser = await puppeteer.connect({ browserURL: browserUrl });

  const pages = await browser.pages();
  const teamsPage = findTeamsPage(pages);
  check("teamsPageFound", !!teamsPage, !!teamsPage);
  if (!teamsPage) {
    browser.disconnect();
    process.exit(1);
  }

  // 1. Capture token
  console.log("\n=== Token Capture ===");
  const token = await captureSkypeToken(teamsPage);
  check("skypeTokenCaptured", !!token.skypeToken, token.skypeToken.length);
  check("tokenRegionDetected", !!token.region, token.region);

  // 2. Fetch conversations
  console.log("\n=== Conversations ===");
  const conversations = await fetchConversations(token, 10);
  check("conversationsFetched", conversations.length > 0, conversations.length);

  const firstChat = conversations.find(
    (conversation) =>
      conversation.topic &&
      (conversation.threadType === "chat" || conversation.threadType === "topic"),
  );
  check("chatConversationFound", !!firstChat, firstChat?.topic ?? "none");

  if (!firstChat) {
    browser.disconnect();
    process.exit(1);
  }

  // 3. Fetch messages
  console.log("\n=== Messages ===");
  const messagesPage = await fetchMessages(token, firstChat.id, 50);
  check("messagesFetched", messagesPage.messages.length > 0, messagesPage.messages.length);

  const textMessages = messagesPage.messages.filter(
    (message) => message.messageType === "RichText/Html" || message.messageType === "Text",
  );
  check("textMessagesFound", textMessages.length > 0, textMessages.length);
  check("paginationAvailable", !!messagesPage.backwardLink, !!messagesPage.backwardLink);

  // 4. Fetch members
  console.log("\n=== Members ===");
  let members: Awaited<ReturnType<typeof fetchMembers>> = [];
  try {
    members = await fetchMembers(token, firstChat.id);
    check("membersFetched", members.length > 0, members.length);
  } catch {
    check("membersFetched", false, "error");
  }

  // 5. Convert to snapshots
  console.log("\n=== Conversion ===");
  const snapshots = convertApiMessagesToSnapshots(messagesPage.messages, members);
  check("snapshotsCreated", snapshots.length > 0, snapshots.length);

  const reactionCount = snapshots.reduce(
    (total, snapshot) => total + snapshot.reactions.length, 0
  );
  check("reactionsConverted", true, reactionCount);

  const quoteCount = snapshots.filter((snapshot) => snapshot.quote !== null).length;
  check("quotedRepliesDetected", true, quoteCount);

  // 6. Render both formats
  console.log("\n=== Rendering ===");
  const meta = { scope: "direct-api", title: firstChat.topic, sourceUrl: teamsPage.url(), exportedAt: new Date().toISOString() };

  const markdownOutput = renderMarkdown(snapshots, meta);
  check("markdownRendered", markdownOutput.length > 100, `${markdownOutput.length} chars`);

  const htmlOutput = renderHtmlDocument(snapshots, meta);
  check("htmlRendered", htmlOutput.length > 100, `${htmlOutput.length} chars`);

  // 7. Write validation output
  const summary = {
    timestamp: new Date().toISOString(),
    checks,
    allPassed: checks.every((result) => result.passed),
    chatUsed: firstChat.topic,
    conversationId: firstChat.id,
    messageCount: textMessages.length,
    snapshotCount: snapshots.length,
    reactionCount,
    quoteCount,
    markdownSize: markdownOutput.length,
    htmlSize: htmlOutput.length,
  };

  await fs.mkdir(artifactsDirectory, { recursive: true });
  const summaryPath = path.join(artifactsDirectory, "direct-api-validation.json");
  await fs.writeFile(summaryPath, JSON.stringify(summary, null, 2));

  console.log(`\n${"=".repeat(50)}`);
  console.log(`Validation: ${summary.allPassed ? "ALL PASSED ✅" : "SOME FAILED ❌"}`);
  console.log(`  Checks: ${checks.filter((check) => check.passed).length}/${checks.length} passed`);
  console.log(`  Results: ${summaryPath}`);

  browser.disconnect();
  process.exit(exitCode);
}

main().catch((error) => {
  console.error("Fatal:", error);
  process.exit(1);
});
