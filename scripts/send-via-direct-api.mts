/**
 * Send a Teams Message via Direct API
 *
 * Acquires an auth token (either via auto-login or connecting to a
 * running Chrome instance), finds or specifies the target conversation,
 * and sends a message.
 *
 * Usage:
 *   # Auto-login mode (recommended):
 *   npx -y tsx scripts/send-via-direct-api.mts --auto --email maxim.mazurok@wisetechglobal.com \
 *     --to "Maxim Mazurok" --message "Hello from the API!"
 *
 *   # Connect to existing Chrome:
 *   DEBUG_PORT=9223 npx -y tsx scripts/send-via-direct-api.mts \
 *     --to "Maxim Mazurok" --message "Hello from the API!"
 *
 *   # Send to a conversation by topic name:
 *   npx -y tsx scripts/send-via-direct-api.mts --auto --email maxim.mazurok@wisetechglobal.com \
 *     --chat "Front End Community" --message "Hello team!"
 *
 * Options:
 *   --to <name|email>          Find 1:1 conversation with this person
 *   --chat <name>              Find conversation by topic (group chats, channels)
 *   --message <text>           Message content to send (plain text)
 *   --auto                     Auto-acquire token (launches fresh Chrome, uses Intune passkey)
 *   --email <email>            Corporate email for auto login (required with --auto)
 *   --list                     List conversations and exit
 */

import puppeteer from "puppeteer-core";

import type { TeamsAuthToken } from "../src/direct-api-client.js";
import {
  captureSkypeToken,
  findTeamsPage,
  fetchConversations,
  fetchCurrentUserDisplayName,
  findOneOnOneConversation,
  sendMessage,
} from "../src/direct-api-client.js";
import { acquireTokenAutomatically } from "../src/auto-token.js";

const browserUrl = `http://127.0.0.1:${Number(process.env.DEBUG_PORT || "9222")}`;

interface CliOptions {
  recipientFilter: string | null;
  chatFilter: string | null;
  messageContent: string | null;
  listOnly: boolean;
  autoLogin: boolean;
  email: string | null;
}

function parseArguments(): CliOptions {
  const cliArguments = process.argv.slice(2);
  const options: CliOptions = {
    recipientFilter: null,
    chatFilter: null,
    messageContent: null,
    listOnly: false,
    autoLogin: false,
    email: null,
  };

  for (let i = 0; i < cliArguments.length; i++) {
    switch (cliArguments[i]) {
      case "--to":
        options.recipientFilter = cliArguments[++i];
        break;
      case "--chat":
        options.chatFilter = cliArguments[++i];
        break;
      case "--message":
        options.messageContent = cliArguments[++i];
        break;
      case "--list":
        options.listOnly = true;
        break;
      case "--auto":
        options.autoLogin = true;
        break;
      case "--email":
        options.email = cliArguments[++i];
        break;
    }
  }

  return options;
}

async function acquireToken(
  options: CliOptions,
): Promise<{ token: TeamsAuthToken; disconnect: () => void }> {
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
    console.log(
      `Token captured (${token.skypeToken.length} chars, region: ${token.region})`,
    );
    return { token, disconnect: () => {} };
  }

  console.log(`Connecting to Chrome at ${browserUrl}...`);
  const browser = await puppeteer.connect({ browserURL: browserUrl });

  const pages = await browser.pages();
  const teamsPage = findTeamsPage(pages);
  if (!teamsPage) {
    console.error(
      "No Teams page found. Navigate to Teams in the browser first.",
    );
    process.exit(1);
  }
  console.log(`Teams page: ${teamsPage.url()}`);

  console.log("Capturing auth token...");
  const token = await captureSkypeToken(teamsPage);
  console.log(
    `Token captured (${token.skypeToken.length} chars, region: ${token.region})`,
  );
  return { token, disconnect: () => browser.disconnect() };
}

async function main() {
  const options = parseArguments();

  const { token, disconnect } = await acquireToken(options);

  try {
    // List mode
    if (options.listOnly) {
      console.log("\nFetching conversations...");
      const conversations = await fetchConversations(token, 100);
      const displayableConversations = conversations.filter(
        (conversation) =>
          conversation.topic &&
          ![
            "streamofannotations",
            "streamofthreads",
            "streamofnotifications",
            "streamofmentions",
            "streamofnotes",
          ].includes(conversation.threadType),
      );
      console.log(`\n${displayableConversations.length} conversations:\n`);
      for (let i = 0; i < displayableConversations.length; i++) {
        const conversation = displayableConversations[i];
        const lastMessage =
          conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
        console.log(
          `  [${i}] ${conversation.threadType}: "${conversation.topic}" (members: ${conversation.memberCount ?? "?"}, last: ${lastMessage})`,
        );
      }
      return;
    }

    // Validate required options
    if (!options.messageContent) {
      console.error("--message is required. Specify the message text to send.");
      process.exit(1);
    }

    if (!options.recipientFilter && !options.chatFilter) {
      console.error(
        "Either --to <name> or --chat <name> is required to identify the target conversation.",
      );
      process.exit(1);
    }

    // Find the target conversation
    let conversationId: string;
    let conversationLabel: string;

    if (options.recipientFilter) {
      console.log(
        `\nSearching for 1:1 conversation with "${options.recipientFilter}"...`,
      );
      const result = await findOneOnOneConversation(
        token,
        options.recipientFilter,
      );
      if (!result) {
        console.error(
          `No 1:1 conversation found matching "${options.recipientFilter}".`,
        );
        console.log(
          'Use --list to see available conversations, or use --chat <name> instead.',
        );
        process.exit(1);
      }
      conversationId = result.conversationId;
      conversationLabel = result.memberDisplayName;
      console.log(`  Found: "${conversationLabel}"`);
    } else {
      console.log(
        `\nSearching for conversation matching "${options.chatFilter}"...`,
      );
      const conversations = await fetchConversations(token, 100);
      const filterLower = options.chatFilter!.toLowerCase();
      const match = conversations.find((conversation) =>
        conversation.topic?.toLowerCase().includes(filterLower),
      );
      if (!match) {
        console.error(
          `No conversation matching "${options.chatFilter}" found.`,
        );
        console.log("Use --list to see available conversations.");
        process.exit(1);
      }
      conversationId = match.id;
      conversationLabel = match.topic;
      console.log(`  Found: "${conversationLabel}"`);
    }

    // Get sender display name
    console.log("\nResolving sender identity...");
    const senderDisplayName = await fetchCurrentUserDisplayName(token);
    console.log(`  Sending as: "${senderDisplayName}"`);

    // Send the message
    console.log(`\nSending message to "${conversationLabel}"...`);
    console.log(`  Content: "${options.messageContent}"`);
    const messageId = await sendMessage(
      token,
      conversationId,
      options.messageContent,
      senderDisplayName,
    );
    console.log(`\n✅ Message sent successfully!`);
    console.log(`  Message ID: ${messageId}`);
    console.log(`  Conversation: ${conversationLabel}`);
  } finally {
    disconnect();
  }
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
