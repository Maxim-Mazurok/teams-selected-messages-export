/**
 * Export Teams Chat Messages via Direct API
 *
 * Connects to a running Chrome instance with Teams open, extracts
 * auth tokens, fetches messages directly via the Teams REST API,
 * and exports them as HTML or Markdown files.
 *
 * This bypasses the Teams UI entirely — no scrolling, no DOM scraping,
 * no web worker interception. Just direct HTTP calls.
 *
 * Usage:
 *   1. Start Chrome with remote debugging:
 *        ./scripts/start-chrome-debug.sh
 *   2. Navigate to Teams and log in
 *   3. Run:
 *        DEBUG_PORT=9223 npx -y tsx scripts/export-via-direct-api.mts [options]
 *
 * Options:
 *   --format html|md         Export format (default: md)
 *   --chat <name>            Partial match on conversation topic
 *   --list                   List conversations and exit
 *   --max-pages <n>          Max pagination pages (default: 100)
 *   --output <dir>           Output directory (default: artifacts/exports)
 *   --auto                   Auto-acquire token (launches fresh Chrome, uses Intune passkey)
 *   --email <email>           Corporate email for auto login (required with --auto)
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

import type { TeamsAuthToken } from "../src/direct-api-client.js";
import {
  captureSkypeToken,
  findTeamsPage,
  fetchConversations,
  fetchAllMessages,
  fetchMembers,
} from "../src/direct-api-client.js";
import { acquireTokenAutomatically } from "../src/auto-token.js";
import { convertApiMessagesToSnapshots } from "../src/direct-api-converter.js";
import { renderMarkdown } from "../src/markdown-renderer.js";
import { renderHtmlDocument } from "../src/html-renderer.js";

const currentFilename = fileURLToPath(import.meta.url);
const currentDirectory = path.dirname(currentFilename);
const rootDirectory = path.resolve(currentDirectory, "..");
const defaultOutputDirectory = path.join(rootDirectory, "artifacts", "exports");

const browserUrl = `http://127.0.0.1:${Number(process.env.DEBUG_PORT || "9222")}`;

interface CliOptions {
  format: "html" | "md";
  chatFilter: string | null;
  listOnly: boolean;
  maxPages: number;
  outputDirectory: string;
  autoLogin: boolean;
  email: string | null;
}

function parseArguments(): CliOptions {
  const arguments_ = process.argv.slice(2);
  const options: CliOptions = {
    format: "md",
    chatFilter: null,
    listOnly: false,
    maxPages: 100,
    outputDirectory: defaultOutputDirectory,
    autoLogin: false,
    email: null,
  };

  for (let i = 0; i < arguments_.length; i++) {
    switch (arguments_[i]) {
      case "--format":
        options.format = arguments_[++i] as "html" | "md";
        break;
      case "--chat":
        options.chatFilter = arguments_[++i];
        break;
      case "--list":
        options.listOnly = true;
        break;
      case "--max-pages":
        options.maxPages = Number(arguments_[++i]);
        break;
      case "--output":
        options.outputDirectory = arguments_[++i];
        break;
      case "--auto":
        options.autoLogin = true;
        break;
      case "--email":
        options.email = arguments_[++i];
        break;
    }
  }

  return options;
}

function sanitizeFilename(name: string): string {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 60);
}

function formatTimestamp(): string {
  return new Date().toISOString().replace(/[:.]/g, "-").slice(0, 23);
}

async function acquireToken(options: CliOptions): Promise<{ token: TeamsAuthToken; sourceUrl: string; disconnect: () => void }> {
  if (options.autoLogin) {
    if (!options.email) {
      console.error("--email is required when using --auto");
      process.exit(1);
    }
    console.log("Auto-acquiring token via Intune passkey...");
    const token = await acquireTokenAutomatically({
      email: options.email,
      headless: true,
      verbose: true,
    });
    console.log(`Token captured (${token.skypeToken.length} chars, region: ${token.region})`);
    return { token, sourceUrl: "https://teams.cloud.microsoft/", disconnect: () => {} };
  }

  console.log(`Connecting to Chrome at ${browserUrl}...`);
  const browser = await puppeteer.connect({ browserURL: browserUrl });

  const pages = await browser.pages();
  const teamsPage = findTeamsPage(pages);
  if (!teamsPage) {
    console.error("No Teams page found. Navigate to Teams in the browser first.");
    process.exit(1);
  }
  console.log(`Teams page: ${teamsPage.url()}`);

  console.log("Capturing auth token...");
  const token = await captureSkypeToken(teamsPage);
  console.log(`Token captured (${token.skypeToken.length} chars, region: ${token.region})`);
  return { token, sourceUrl: teamsPage.url(), disconnect: () => browser.disconnect() };
}

async function main() {
  const options = parseArguments();

  const { token, sourceUrl, disconnect } = await acquireToken(options);

  // Fetch conversations
  console.log("\nFetching conversations...");
  const conversations = await fetchConversations(token, 50);

  // Filter to chats with topics (exclude system streams)
  const displayableConversations = conversations.filter(
    (conversation) =>
      conversation.topic &&
      !["streamofannotations", "streamofthreads", "streamofnotifications", "streamofmentions", "streamofnotes"].includes(conversation.threadType),
  );

  if (options.listOnly) {
    console.log(`\n${displayableConversations.length} conversations:\n`);
    for (let i = 0; i < displayableConversations.length; i++) {
      const conversation = displayableConversations[i];
      const lastMessage = conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
      console.log(
        `  [${i}] ${conversation.threadType}: "${conversation.topic}" (last: ${lastMessage})`,
      );
    }
    disconnect();
    return;
  }

  // Find the target conversation
  let targetConversation = displayableConversations[0];
  if (options.chatFilter) {
    const filter = options.chatFilter.toLowerCase();
    const match = displayableConversations.find((conversation) =>
      conversation.topic.toLowerCase().includes(filter),
    );
    if (!match) {
      console.error(`No conversation matching "${options.chatFilter}" found.`);
      console.log("Available conversations:");
      for (const conversation of displayableConversations.slice(0, 20)) {
        console.log(`  - "${conversation.topic}"`);
      }
      disconnect();
      process.exit(1);
    }
    targetConversation = match;
  }

  console.log(`\nExporting: "${targetConversation.topic}"`);
  console.log(`  ID: ${targetConversation.id}`);
  console.log(`  Type: ${targetConversation.threadType}`);

  // Fetch members for reaction user resolution
  console.log("\nFetching members...");
  let members: Awaited<ReturnType<typeof fetchMembers>> = [];
  try {
    members = await fetchMembers(token, targetConversation.id);
    console.log(`  ${members.length} members found`);
  } catch (error) {
    console.log(`  Members fetch failed: ${(error as Error).message}`);
    members = [];
  }

  // Fetch all messages
  console.log("\nFetching messages...");
  const apiMessages = await fetchAllMessages(token, targetConversation.id, {
    maxPages: options.maxPages,
    pageSize: 200,
    onProgress: (count) => {
      process.stdout.write(`\r  ${count} messages fetched...`);
    },
  });
  console.log(`\n  Total: ${apiMessages.length} messages`);

  // Filter to text messages
  const textMessages = apiMessages.filter(
    (message) =>
      (message.messageType === "RichText/Html" || message.messageType === "Text") &&
      !message.isDeleted,
  );
  console.log(`  Text messages: ${textMessages.length}`);

  // Convert to MessageSnapshot
  console.log("\nConverting to export format...");
  const snapshots = convertApiMessagesToSnapshots(apiMessages, members);
  console.log(`  Snapshots: ${snapshots.length}`);

  // Render
  const meta = {
    scope: "direct-api",
    title: targetConversation.topic,
    sourceUrl,
    exportedAt: new Date().toISOString(),
  };

  const content =
    options.format === "html"
      ? renderHtmlDocument(snapshots, meta)
      : renderMarkdown(snapshots, meta);

  // Write output
  const filename = `${sanitizeFilename(targetConversation.topic)}-direct-api-${formatTimestamp()}.${options.format}`;
  await fs.mkdir(options.outputDirectory, { recursive: true });
  const outputPath = path.join(options.outputDirectory, filename);
  await fs.writeFile(outputPath, content, "utf-8");

  console.log(`\nExport complete:`);
  console.log(`  File: ${outputPath}`);
  console.log(`  Format: ${options.format}`);
  console.log(`  Messages: ${snapshots.length}`);
  console.log(`  Size: ${Math.round(content.length / 1024)} KB`);

  disconnect();
}

main().catch((error) => {
  console.error("Fatal:", error);
  process.exit(1);
});
