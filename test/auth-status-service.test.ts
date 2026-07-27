import assert from "node:assert/strict";
import test from "node:test";

import { AuthStatusService } from "../src/sharepoint/auth-status-service.js";
import type {
  SharePointResponse,
  SharePointTransport,
} from "../src/sharepoint/http.js";

const SITE_URL = "https://example.sharepoint.com/sites/genai";

class FakeTransport implements SharePointTransport {
  public primaryCalls = 0;
  public fallbackCalls = 0;

  constructor(
    private readonly primary: SharePointResponse | Error,
    private readonly fallback: SharePointResponse | Error,
  ) {}

  async get(): Promise<SharePointResponse> {
    this.primaryCalls += 1;
    if (this.primary instanceof Error) {
      throw this.primary;
    }
    return this.primary;
  }

  async getViaPage(): Promise<SharePointResponse> {
    this.fallbackCalls += 1;
    if (this.fallback instanceof Error) {
      throw this.fallback;
    }
    return this.fallback;
  }

  async close(): Promise<void> {}
}

function response(
  status: number,
  body = "",
  contentType = "application/json;odata=nometadata",
  method: "api-request" | "browser-fetch" = "api-request",
): SharePointResponse {
  return { status, body, contentType, method };
}

test("returns authenticated user details from the primary request", async () => {
  const transport = new FakeTransport(
    response(200, JSON.stringify({ Id: 12, Title: "Example User", IsSiteAdmin: false })),
    new Error("fallback must not run"),
  );
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "AUTHENTICATED");
  assert.equal(result.authenticated, true);
  assert.equal(result.user?.displayName, "Example User");
  assert.equal(transport.fallbackCalls, 0);
});

test("uses browser fetch when the primary response is a login page", async () => {
  const transport = new FakeTransport(
    response(200, "<html>login</html>", "text/html"),
    response(
      200,
      JSON.stringify({ Id: 12, Title: "Example User", IsSiteAdmin: false }),
      "application/json",
      "browser-fetch",
    ),
  );
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "AUTHENTICATED");
  assert.equal(result.method, "browser-fetch");
  assert.equal(transport.fallbackCalls, 1);
});

test("uses browser fetch after a primary transport error", async () => {
  const transport = new FakeTransport(
    new Error("proxy rejected API request"),
    response(
      200,
      JSON.stringify({ Id: 12, Title: "Example User", IsSiteAdmin: false }),
      "application/json",
      "browser-fetch",
    ),
  );
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "AUTHENTICATED");
  assert.equal(transport.fallbackCalls, 1);
});

test("does not retry a clear access denial", async () => {
  const transport = new FakeTransport(
    response(403),
    new Error("fallback must not run"),
  );
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "ACCESS_DENIED");
  assert.equal(transport.fallbackCalls, 0);
});

test("maps a browser login redirect to LOGIN_REQUIRED", async () => {
  const transport = new FakeTransport(
    response(302, "", "text/html"),
    response(401, "", "text/plain", "browser-fetch"),
  );
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "LOGIN_REQUIRED");
  assert.equal(result.authenticated, false);
});

test("maps HTTP 404 to SITE_NOT_FOUND", async () => {
  const transport = new FakeTransport(
    response(404),
    new Error("fallback must not run"),
  );
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "SITE_NOT_FOUND");
  assert.equal(transport.fallbackCalls, 0);
});

test("returns UNAVAILABLE when both request methods fail", async () => {
  const transport = new FakeTransport(new Error("primary"), new Error("fallback"));
  const result = await new AuthStatusService(SITE_URL, transport).getStatus();

  assert.equal(result.state, "UNAVAILABLE");
  assert.equal(result.authenticated, false);
});
