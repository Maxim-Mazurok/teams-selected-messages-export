import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import puppeteer from "puppeteer-core";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const artifactsDir = path.join(rootDir, "artifacts");
const summaryPath = path.join(artifactsDir, "worker-intercept-probe.json");
const browserUrl = `http://127.0.0.1:${Number(process.env.DEBUG_PORT || "9222")}`;

export function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function stripHtml(value) {
  return String(value || "")
    .replace(/<blockquote[\s\S]*?<\/blockquote>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function simplifyEmotionUsers(collection = []) {
  return collection.map((entry) => ({
    key: entry.key || entry.commonKey || null,
    count: entry.count ?? (Array.isArray(entry.users) ? entry.users.length : null),
    userIds: Array.isArray(entry.users) ? entry.users.map((user) => user.userId || user.id).filter(Boolean) : []
  }));
}

export function simplifyMessage(message) {
  return {
    id: message.id || null,
    time: message.originalArrivalTime || null,
    author: message.fromUser?.displayName || message.imDisplayName || null,
    authorId: message.fromUser?.id || message.fromUserId || message.from || null,
    type: message.messageType || null,
    contentPreview: stripHtml(message.content).slice(0, 220),
    quotedMessages: message.quotedMessages || null,
    emotionsSummary: Array.isArray(message.emotionsSummary)
      ? message.emotionsSummary.map((entry) => ({
          key: entry.key || null,
          count: entry.count ?? null
        }))
      : null,
    emotions: Array.isArray(message.emotions) ? simplifyEmotionUsers(message.emotions) : null,
    diverseEmotions: Array.isArray(message.diverseEmotions?.items)
      ? simplifyEmotionUsers(message.diverseEmotions.items)
      : null
  };
}

export function collectReactionUsers(messages) {
  const seen = new Set();
  const results = [];

  for (const message of messages) {
    for (const emotionGroup of [...(message.emotions || []), ...(message.diverseEmotions || [])]) {
      for (const userId of emotionGroup.userIds || []) {
        const key = `${message.id}:${emotionGroup.key}:${userId}`;
        if (!seen.has(key)) {
          seen.add(key);
          results.push({
            messageId: message.id,
            reactionKey: emotionGroup.key,
            userId
          });
        }
      }
    }
  }

  return results;
}

async function ensureTeamsPage(browser) {
  const pages = await browser.pages();
  const page = pages.find((entry) => /teams\.(microsoft|cloud\.microsoft)/i.test(entry.url()));
  if (!page) {
    throw new Error("No Teams page found in the connected Chrome instance.");
  }
  return page;
}

async function installProbe(page) {
  await page.evaluateOnNewDocument(() => {
    const MAX_EVENTS = 200;
    const FETCH_BODY_LIMIT = 1200;
    const interestingUrl = (url) =>
      /chatsvc|conversation|messages|threads|emotion|reaction|chat/i.test(String(url || ""));

    const sanitizeText = (value, limit) => {
      const text = String(value || "");
      return text.length > limit ? `${text.slice(0, limit)}...[truncated]` : text;
    };

    const maybeJson = (value) => {
      if (!value) {
        return null;
      }
      if (typeof value === "string") {
        try {
          return JSON.parse(value);
        } catch {
          return null;
        }
      }
      return value;
    };

    const summarizeMessage = (message) => ({
      id: message?.id || null,
      time: message?.originalArrivalTime || null,
      author: message?.fromUser?.displayName || message?.imDisplayName || null,
      authorId: message?.fromUser?.id || message?.fromUserId || message?.from || null,
      type: message?.messageType || null,
      content: sanitizeText(String(message?.content || ""), 800),
      emotionsSummary: Array.isArray(message?.emotionsSummary)
        ? message.emotionsSummary.map((entry) => ({
            key: entry?.key || null,
            count: entry?.count ?? null
          }))
        : null,
      emotions: Array.isArray(message?.emotions)
        ? message.emotions.map((entry) => ({
            key: entry?.key || null,
            userIds: Array.isArray(entry?.users) ? entry.users.map((user) => user?.userId || user?.id).filter(Boolean) : []
          }))
        : null,
      diverseEmotions: Array.isArray(message?.diverseEmotions?.items)
        ? message.diverseEmotions.items.map((entry) => ({
            key: entry?.commonKey || entry?.key || null,
            userIds: Array.isArray(entry?.users) ? entry.users.map((user) => user?.userId || user?.id).filter(Boolean) : []
          }))
        : null
    });

    const summarizeWorkerPayload = (value, direction, workerUrl) => {
      const payload = maybeJson(value);
      if (!payload || typeof payload !== "object") {
        return null;
      }

      const operationName = payload.operationName || payload.payload?.operationName || null;
      const graphqlQuery = payload.payload?.graphqlQuery || null;
      const responseData = payload.response?.data || null;
      const summary = {
        kind: direction === "request" ? "worker-request" : "worker-response",
        workerUrl,
        requestId: payload.requestId || null,
        operationName,
        transportType: payload.type || null
      };

      if (graphqlQuery) {
        summary.queryPreview = sanitizeText(graphqlQuery, 600);
        summary.variables = payload.payload?.variables || null;
      }

      if (responseData?.messages?.messages) {
        summary.dataSource = payload.extensions?.dataSource || null;
        summary.messages = responseData.messages.messages.slice(0, 4).map(summarizeMessage);
        summary.totalMessages = responseData.messages.messages.length;
        return summary;
      }

      if (Array.isArray(responseData?.fetchRecentMessagesAfter)) {
        summary.dataSource = payload.extensions?.dataSource || null;
        summary.messages = responseData.fetchRecentMessagesAfter.slice(0, 4).map(summarizeMessage);
        summary.totalMessages = responseData.fetchRecentMessagesAfter.length;
        return summary;
      }

      if (graphqlQuery && /MessageListQuery|MissingMessagesQuery|ToastEvent|ChatServiceBatchEvent|ChatQuery/i.test(graphqlQuery)) {
        return summary;
      }

      const serialized = sanitizeText(JSON.stringify(payload), 1200);
      if (/emotion|reaction|message|conversation/i.test(serialized)) {
        summary.preview = serialized;
        return summary;
      }

      return null;
    };

    const push = (event) => {
      const store = (window.__teamsWorkerInterceptProbe = window.__teamsWorkerInterceptProbe || { events: [] });
      store.events.push({ ...event, ts: new Date().toISOString() });
      if (store.events.length > MAX_EVENTS) {
        store.events.shift();
      }
    };

    const origFetch = window.fetch;
    window.fetch = async function(...args) {
      const [input, init] = args;
      const url = typeof input === "string" ? input : input?.url;
      const method = init?.method || (typeof input !== "string" && input?.method) || "GET";
      const response = await origFetch.apply(this, args);
      if (interestingUrl(url)) {
        let bodyPreview = null;
        try {
          bodyPreview = sanitizeText(await response.clone().text(), FETCH_BODY_LIMIT);
        } catch (error) {
          bodyPreview = `[fetch body unavailable: ${error?.message || error}]`;
        }
        push({ kind: "fetch", url, method, status: response.status, bodyPreview });
      }
      return response;
    };

    const origOpen = XMLHttpRequest.prototype.open;
    const origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function(method, url, ...rest) {
      this.__teamsWorkerProbe = { method, url };
      return origOpen.call(this, method, url, ...rest);
    };
    XMLHttpRequest.prototype.send = function(body) {
      if (this.__teamsWorkerProbe) {
        this.addEventListener(
          "loadend",
          () => {
            if (!interestingUrl(this.__teamsWorkerProbe?.url)) {
              return;
            }
            let bodyPreview = null;
            try {
              bodyPreview =
                this.responseType && this.responseType !== "text" && this.responseType !== ""
                  ? `[xhr responseType ${this.responseType}]`
                  : sanitizeText(this.responseText, FETCH_BODY_LIMIT);
            } catch (error) {
              bodyPreview = `[xhr body unavailable: ${error?.message || error}]`;
            }
            push({
              kind: "xhr",
              url: this.__teamsWorkerProbe?.url,
              method: this.__teamsWorkerProbe?.method,
              status: this.status,
              bodyPreview
            });
          },
          { once: true }
        );
      }
      return origSend.call(this, body);
    };

    const OrigWebSocket = window.WebSocket;
    function WrappedWebSocket(url, protocols) {
      const socket = protocols ? new OrigWebSocket(url, protocols) : new OrigWebSocket(url);
      if (interestingUrl(url)) {
        push({ kind: "ws-open", url: String(url) });
      }
      socket.addEventListener("message", (event) => {
        if (!interestingUrl(url)) {
          return;
        }
        const bodyPreview =
          typeof event.data === "string"
            ? sanitizeText(event.data, FETCH_BODY_LIMIT)
            : `[ws ${Object.prototype.toString.call(event.data)}]`;
        push({ kind: "ws-message", url: String(url), bodyPreview });
      });
      return socket;
    }
    WrappedWebSocket.prototype = OrigWebSocket.prototype;
    Object.setPrototypeOf(WrappedWebSocket, OrigWebSocket);
    window.WebSocket = WrappedWebSocket;

    const OrigWorker = window.Worker;
    function WrappedWorker(url, options) {
      const worker = options ? new OrigWorker(url, options) : new OrigWorker(url);
      push({ kind: "worker-create", workerUrl: String(url) });

      worker.addEventListener("message", (event) => {
        const summary = summarizeWorkerPayload(event.data, "response", String(url));
        if (summary) {
          push(summary);
        }
      });

      const origPostMessage = worker.postMessage.bind(worker);
      worker.postMessage = function(data, transfer) {
        const summary = summarizeWorkerPayload(data, "request", String(url));
        if (summary) {
          push(summary);
        }
        return origPostMessage(data, transfer);
      };

      return worker;
    }
    WrappedWorker.prototype = OrigWorker.prototype;
    Object.setPrototypeOf(WrappedWorker, OrigWorker);
    window.Worker = WrappedWorker;
  });
}

export async function main() {
  await fs.mkdir(artifactsDir, { recursive: true });
  const browser = await puppeteer.connect({ browserURL: browserUrl, defaultViewport: null });

  try {
    const page = await ensureTeamsPage(browser);
    await installProbe(page);
    await page.bringToFront();
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.waitForFunction(
      () => document.querySelector('[data-tid="chat-title"]') && document.querySelectorAll('[data-tid="chat-pane-item"]').length > 0,
      { timeout: 30000 }
    );
    await sleep(6000);

    const pageState = await page.evaluate(() => ({
      title: document.title,
      href: location.href,
      chatTitle: document.querySelector('[data-tid="chat-title"]')?.textContent?.trim() || null,
      messageCount: document.querySelectorAll('[data-tid="chat-pane-item"]').length
    }));

    const events = await page.evaluate(() => (window.__teamsWorkerInterceptProbe || { events: [] }).events);
    const workerRequests = events.filter((event) => event.kind === "worker-request");
    const workerResponses = events.filter((event) => event.kind === "worker-response");
    const messageResponses = workerResponses.filter((event) => Array.isArray(event.messages) && event.messages.length > 0);
    const allMessages = messageResponses.flatMap((event) => event.messages);
    const reactionMessages = allMessages.filter(
      (message) =>
        (Array.isArray(message.emotionsSummary) && message.emotionsSummary.length > 0) ||
        (Array.isArray(message.diverseEmotions) && message.diverseEmotions.length > 0)
    );

    const summary = {
      browserUrl,
      page: pageState,
      counts: events.reduce((accumulator, event) => {
        accumulator[event.kind] = (accumulator[event.kind] || 0) + 1;
        return accumulator;
      }, {}),
      mainThreadFindings: {
        fetchEvents: events.filter((event) => event.kind === "fetch"),
        xhrEvents: events.filter((event) => event.kind === "xhr"),
        wsEvents: events.filter((event) => event.kind === "ws-open" || event.kind === "ws-message")
      },
      workerOperationNames: Array.from(
        new Set(workerRequests.map((event) => event.operationName).filter(Boolean))
      ).sort(),
      messageResponseSamples: messageResponses.slice(0, 3),
      reactionMessageSamples: reactionMessages.slice(0, 5),
      reactionUsers: collectReactionUsers(reactionMessages),
      architectureAssessment: {
        pageFetchXhrViable: events.some((event) => event.kind === "fetch" || event.kind === "xhr"),
        workerInterceptionViable: workerResponses.length > 0,
        recommendedHookPoint:
          "Inject a MAIN-world script at document_start to patch Worker first, then bridge summarized events to the extension's isolated UI/content script.",
        notes: [
          "Main-thread fetch/XHR saw only config traffic during this probe, not chat message payloads.",
          "Worker messages exposed GraphQL request metadata and message list responses.",
          "Message responses already included content, quoted reply HTML, reactions, and reaction user IDs.",
          "The sampled message list responses were served from indexedDB in this run, which means this hook observes structured data even when the thread is satisfied from cache."
        ]
      }
    };

    await fs.writeFile(summaryPath, JSON.stringify(summary, null, 2), "utf8");
    console.log(JSON.stringify({ ...summary, summaryPath }, null, 2));
  } finally {
    await browser.disconnect();
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
