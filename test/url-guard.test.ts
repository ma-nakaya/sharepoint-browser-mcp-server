import assert from "node:assert/strict";
import test from "node:test";

import { buildSharePointApiUrl } from "../src/sharepoint/url-guard.js";

const SITE_URL = "https://example.sharepoint.com/sites/genai";

test("builds a REST URL under the configured site", () => {
  assert.equal(
    buildSharePointApiUrl(SITE_URL, "/_api/web/currentuser").toString(),
    "https://example.sharepoint.com/sites/genai/_api/web/currentuser",
  );
});

test("rejects non-REST paths", () => {
  assert.throws(() => buildSharePointApiUrl(SITE_URL, "/SitePages/Home.aspx"), /\/_api\//u);
});

test("does not allow an absolute URL to escape the configured site", () => {
  assert.throws(
    () => buildSharePointApiUrl(SITE_URL, "https://evil.example/_api/web"),
    /\/_api\//u,
  );
});

test("rejects path traversal outside the configured REST root", () => {
  assert.throws(
    () => buildSharePointApiUrl(SITE_URL, "/_api/../../other"),
    /configured site path/u,
  );
});
