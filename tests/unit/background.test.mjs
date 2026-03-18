import { describe, it, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const currentDirectory = dirname(fileURLToPath(import.meta.url));
const backgroundScriptPath = join(currentDirectory, "../../extension-src/background.js");

// We load background.js by eval-ing it with chrome/fetch mocks in scope.
// This validates the message handler logic and fetch flow end-to-end.

/** @type {Array<(message: any, sender: any, sendResponse: (response: any) => void) => boolean | void>} */
let messageListeners;

/** @type {Array<{url: string, options: RequestInit}>} */
let fetchCalls;

/** @type {(url: string, options?: RequestInit) => Promise<Response>} */
let mockFetchImplementation;

function setupMocks() {
  messageListeners = [];
  fetchCalls = [];
  mockFetchImplementation = async () => new Response("{}");

  // Mock chrome.runtime and chrome.action
  globalThis.chrome = /** @type {any} */ ({
    runtime: {
      onMessage: {
        addListener(/** @type {any} */ listener) {
          messageListeners.push(listener);
        }
      }
    },
    action: {
      onClicked: {
        addListener() {}
      }
    },
    tabs: {
      sendMessage: async () => {}
    }
  });

  // Mock fetch
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (url, options) => {
    fetchCalls.push({ url: String(url), options: /** @type {RequestInit} */ (options || {}) });
    return mockFetchImplementation(String(url), options);
  };

  return () => {
    globalThis.fetch = originalFetch;
    delete globalThis.chrome;
  };
}

async function loadBackgroundScript() {
  const source = await readFile(backgroundScriptPath, "utf-8");
  // Wrap in an async IIFE to avoid polluting the global scope across tests.
  // eslint-disable-next-line no-eval
  eval(source);
}

function sendMessage(message) {
  return new Promise((resolve) => {
    for (const listener of messageListeners) {
      listener(message, {}, resolve);
    }
  });
}

function createMockResponse(body, status = 200) {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? "OK" : `Error ${status}`,
    json: async () => body
  };
}

describe("background.js message handler", () => {
  let cleanup;

  beforeEach(async () => {
    cleanup?.();
    cleanup = setupMocks();
    await loadBackgroundScript();
  });

  it("responds with error when token is missing", async () => {
    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.ok(response.error);
    assert.match(response.error, /Missing API credentials/);
  });

  it("responds with error when region is missing", async () => {
    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "token-123",
      conversationId: "19:abc@thread.v2"
    });

    assert.ok(response.error);
    assert.match(response.error, /Missing API credentials/);
  });

  it("responds with error when conversationId is missing", async () => {
    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "token-123",
      region: "apac"
    });

    assert.ok(response.error);
    assert.match(response.error, /Missing conversation ID/);
  });

  it("fetches messages with bearer auth headers", async () => {
    mockFetchImplementation = async () => createMockResponse({
      messages: [{ id: "1", content: "Hello" }],
      _metadata: {}
    });

    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "my-bearer-token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.equal(fetchCalls.length, 1);
    assert.match(fetchCalls[0].url, /apac\.ng\.msg\.teams\.microsoft\.com/);
    assert.match(fetchCalls[0].url, /19%3Aabc%40thread\.v2/);
    assert.equal(fetchCalls[0].options.headers.Authorization, "Bearer my-bearer-token");
    assert.equal(response.messages.length, 1);
    assert.equal(response.pageCount, 1);
  });

  it("fetches messages with skypetoken auth headers", async () => {
    mockFetchImplementation = async () => createMockResponse({
      messages: [{ id: "1", content: "Hello" }],
      _metadata: {}
    });

    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "my-skype-token",
      tokenType: "skypetoken",
      region: "emea",
      conversationId: "19:xyz@thread.v2"
    });

    assert.equal(fetchCalls.length, 1);
    assert.match(fetchCalls[0].url, /emea\.ng\.msg\.teams\.microsoft\.com/);
    assert.equal(fetchCalls[0].options.headers.Authentication, "skypetoken=my-skype-token");
    assert.equal(response.messages.length, 1);
  });

  it("defaults to skypetoken auth when tokenType is not specified", async () => {
    mockFetchImplementation = async () => createMockResponse({
      messages: [],
      _metadata: {}
    });

    await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "my-token",
      region: "amer",
      conversationId: "19:xyz@thread.v2"
    });

    assert.equal(fetchCalls[0].options.headers.Authentication, "skypetoken=my-token");
  });

  it("paginates using backwardLink", async () => {
    let callIndex = 0;
    mockFetchImplementation = async () => {
      callIndex++;
      if (callIndex === 1) {
        return createMockResponse({
          messages: [{ id: "2", content: "Page 1" }],
          _metadata: { backwardLink: "https://apac.ng.msg.teams.microsoft.com/v1/page2" }
        });
      }
      return createMockResponse({
        messages: [{ id: "1", content: "Page 2" }],
        _metadata: {}
      });
    };

    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.equal(fetchCalls.length, 2);
    assert.equal(response.messages.length, 2);
    assert.equal(response.pageCount, 2);
  });

  it("returns auth error on 401", async () => {
    mockFetchImplementation = async () => createMockResponse({}, 401);

    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "expired-token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.ok(response.error);
    assert.match(response.error, /Authentication failed.*401/);
  });

  it("returns auth error on 403", async () => {
    mockFetchImplementation = async () => createMockResponse({}, 403);

    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "bad-token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.ok(response.error);
    assert.match(response.error, /Authentication failed.*403/);
  });

  it("returns error on non-OK status", async () => {
    mockFetchImplementation = async () => createMockResponse({}, 500);

    const response = await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.ok(response.error);
    assert.match(response.error, /API returned 500/);
  });

  it("constructs correct URL with encoded conversation ID", async () => {
    mockFetchImplementation = async () => createMockResponse({ messages: [], _metadata: {} });

    await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:user1_user2@unq.gbl.spaces"
    });

    assert.match(
      fetchCalls[0].url,
      /\/v1\/users\/ME\/conversations\/19%3Auser1_user2%40unq\.gbl\.spaces\/messages\?pageSize=200/
    );
  });

  it("uses page size of 200", async () => {
    mockFetchImplementation = async () => createMockResponse({ messages: [], _metadata: {} });

    await sendMessage({
      type: "teams-export/fetch-full-chat",
      skypeToken: "token",
      tokenType: "bearer",
      region: "apac",
      conversationId: "19:abc@thread.v2"
    });

    assert.match(fetchCalls[0].url, /pageSize=200/);
  });
});
