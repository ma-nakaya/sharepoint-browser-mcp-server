import assert from "node:assert/strict";
import test from "node:test";

import type { AppConfig } from "../src/config.js";
import type { Logger } from "../src/logger.js";
import {
  MAX_DOCUMENT_SOURCE_BYTES,
  MAX_DOWNLOAD_BYTES,
} from "../src/sharepoint/file-content.js";
import { PlaywrightSharePointTransport } from "../src/sharepoint/playwright-transport.js";

const CONFIG: AppConfig = {
  siteUrl: "https://example.sharepoint.com/sites/genai",
  siteOrigin: "https://example.sharepoint.com",
  profileDir: "/tmp/edge-profile",
  browserChannel: "msedge",
  headless: true,
  requestTimeoutMs: 15_000,
};

const LOGGER: Logger = {
  info: () => undefined,
  warn: () => undefined,
  error: () => undefined,
};

test("does not read a redirect response body", async () => {
  let bodyRead = false;
  let requestedUrl = "";
  let requestOptions: Record<string, unknown> | undefined;

  const context = {
    request: {
      get: async (url: string, options: Record<string, unknown>) => {
        requestedUrl = url;
        requestOptions = options;
        return {
          status: () => 302,
          headers: () => ({ location: "https://login.example/" }),
          text: async () => {
            bodyRead = true;
            throw new Error("redirect body must not be read");
          },
        };
      },
    },
  };
  const edgeSession = {
    getContext: async () => context,
    close: async () => undefined,
  };

  const transport = new PlaywrightSharePointTransport(
    CONFIG,
    edgeSession as never,
    LOGGER,
  );
  const result = await transport.get("/_api/web/currentuser");

  assert.equal(result.status, 302);
  assert.equal(result.body, "");
  assert.equal(bodyRead, false);
  assert.equal(
    requestedUrl,
    "https://example.sharepoint.com/sites/genai/_api/web/currentuser",
  );
  assert.equal(requestOptions?.maxRedirects, 0);
  assert.equal(requestOptions?.failOnStatusCode, false);
});

test("reads a successful JSON response body", async () => {
  const jsonBody = JSON.stringify({ Id: 1, Title: "User", IsSiteAdmin: false });
  const context = {
    request: {
      get: async () => ({
        status: () => 200,
        headers: () => ({ "content-type": "application/json;odata=nometadata" }),
        text: async () => jsonBody,
      }),
    },
  };
  const edgeSession = {
    getContext: async () => context,
    close: async () => undefined,
  };

  const transport = new PlaywrightSharePointTransport(
    CONFIG,
    edgeSession as never,
    LOGGER,
  );
  const result = await transport.get("/_api/web/currentuser");

  assert.equal(result.status, 200);
  assert.equal(result.body, jsonBody);
  assert.equal(result.method, "api-request");
});

test("rejects an oversized binary response before reading its body", async () => {
  let bodyRead = false;
  const context = {
    request: {
      get: async () => ({
        status: () => 200,
        headers: () => ({
          "content-type": "application/pdf",
          "content-length": String(MAX_DOWNLOAD_BYTES + 1),
        }),
        body: async () => {
          bodyRead = true;
          return Buffer.alloc(0);
        },
      }),
    },
  };
  const edgeSession = {
    getContext: async () => context,
    close: async () => undefined,
  };
  const transport = new PlaywrightSharePointTransport(
    CONFIG,
    edgeSession as never,
    LOGGER,
  );

  const result = await transport.getBinary(
    "/_api/web/GetFileByServerRelativePath(decodedUrl='/sites/genai/Documents/a.pdf')/$value",
  );

  assert.equal(result.bodyTruncated, true);
  assert.equal(result.body.byteLength, 0);
  assert.equal(bodyRead, false);
});

test("reads a larger document source when given the internal document limit", async () => {
  let bodyRead = false;
  const bytes = Buffer.alloc(MAX_DOWNLOAD_BYTES + 1);
  const context = {
    request: {
      get: async () => ({
        status: () => 200,
        headers: () => ({
          "content-type":
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "content-length": String(bytes.byteLength),
        }),
        body: async () => {
          bodyRead = true;
          return bytes;
        },
      }),
    },
  };
  const edgeSession = {
    getContext: async () => context,
    close: async () => undefined,
  };
  const transport = new PlaywrightSharePointTransport(
    CONFIG,
    edgeSession as never,
    LOGGER,
  );

  const result = await transport.getBinary(
    "/_api/web/GetFileByServerRelativePath(decodedUrl='/sites/genai/Documents/a.pptx')/$value",
    MAX_DOCUMENT_SOURCE_BYTES,
  );

  assert.equal(result.bodyTruncated, false);
  assert.equal(result.body.byteLength, bytes.byteLength);
  assert.equal(bodyRead, true);
});

test("rejects binary response limits above the document source boundary", async () => {
  const edgeSession = {
    getContext: async () => {
      throw new Error("browser must not start");
    },
    close: async () => undefined,
  };
  const transport = new PlaywrightSharePointTransport(
    CONFIG,
    edgeSession as never,
    LOGGER,
  );

  await assert.rejects(
    () =>
      transport.getBinary(
        "/_api/web/GetFileByServerRelativePath(decodedUrl='/sites/genai/Documents/a.pptx')/$value",
        MAX_DOCUMENT_SOURCE_BYTES + 1,
      ),
    /Binary response limit/u,
  );
});

function createLoginRedirectContext(closeCounter: { value: number }) {
  const page = {
    goto: async () => undefined,
    url: () => "https://login.microsoftonline.com/tenant/oauth2/authorize",
    close: async () => {
      closeCounter.value += 1;
    },
  };
  return {
    request: { get: async () => { throw new Error("not used"); } },
    newPage: async () => page,
  };
}

test("keeps a headed login page open for interactive authentication", async () => {
  const closeCounter = { value: 0 };
  const context = createLoginRedirectContext(closeCounter);
  const edgeSession = {
    getContext: async () => context,
    close: async () => undefined,
  };
  const transport = new PlaywrightSharePointTransport(
    { ...CONFIG, headless: false },
    edgeSession as never,
    LOGGER,
  );

  const result = await transport.getViaPage("/_api/web/currentuser");

  assert.equal(result.status, 401);
  assert.equal(closeCounter.value, 0);
});

test("closes a headless login page after detecting authentication is required", async () => {
  const closeCounter = { value: 0 };
  const context = createLoginRedirectContext(closeCounter);
  const edgeSession = {
    getContext: async () => context,
    close: async () => undefined,
  };
  const transport = new PlaywrightSharePointTransport(
    CONFIG,
    edgeSession as never,
    LOGGER,
  );

  const result = await transport.getViaPage("/_api/web/currentuser");

  assert.equal(result.status, 401);
  assert.equal(closeCounter.value, 1);
});
