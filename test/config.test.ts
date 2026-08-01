import assert from "node:assert/strict";
import test from "node:test";
import path from "node:path";

import {
  assertSafeProfileDirectory,
  loadConfig,
  normalizeSharePointSiteUrl,
  normalizeSharePointSiteUrls,
} from "../src/config.js";

test("normalizes a site collection URL", () => {
  assert.equal(
    normalizeSharePointSiteUrl("https://Example.sharepoint.com/sites/GenAI/"),
    "https://example.sharepoint.com/sites/GenAI",
  );
});

test("rejects tenant root and managed-path roots", () => {
  assert.throws(
    () => normalizeSharePointSiteUrl("https://example.sharepoint.com"),
    /site collection/u,
  );
  assert.throws(
    () => normalizeSharePointSiteUrl("https://example.sharepoint.com/sites"),
    /site collection/u,
  );
});

test("rejects OneDrive and non-SharePoint hosts", () => {
  assert.throws(
    () => normalizeSharePointSiteUrl("https://example-my.sharepoint.com/personal/user"),
    /OneDrive/u,
  );
  assert.throws(
    () => normalizeSharePointSiteUrl("https://example.com/sites/genai"),
    /SharePoint Online/u,
  );
});

test("loads secure defaults", () => {
  const config = loadConfig({
    SHAREPOINT_SITE_URL: "https://example.sharepoint.com/sites/genai",
    LOCALAPPDATA: path.join(path.sep, "tmp", "localappdata"),
  });

  assert.equal(config.browserChannel, "msedge");
  assert.equal(config.headless, true);
  assert.equal(config.requestTimeoutMs, 15_000);
  assert.deepEqual(config.siteUrls, [
    "https://example.sharepoint.com/sites/genai",
  ]);
  assert.match(config.profileDir, /sharepoint-browser-mcp-server/u);
});

test("loads additional SharePoint sites from the same tenant", () => {
  const config = loadConfig({
    SHAREPOINT_SITE_URL: "https://example.sharepoint.com/teams/ymsl",
    SHAREPOINT_ADDITIONAL_SITE_URLS: [
      "https://example.sharepoint.com/teams/ymsl_softeng",
      "https://example.sharepoint.com/teams/ymsl",
    ].join(","),
    LOCALAPPDATA: path.join(path.sep, "tmp", "localappdata"),
  });

  assert.deepEqual(config.siteUrls, [
    "https://example.sharepoint.com/teams/ymsl",
    "https://example.sharepoint.com/teams/ymsl_softeng",
  ]);
});

test("rejects additional sites from another tenant", () => {
  assert.throws(
    () =>
      normalizeSharePointSiteUrls(
        "https://example.sharepoint.com/teams/ymsl",
        "https://other.sharepoint.com/teams/softeng",
      ),
    /same SharePoint tenant/u,
  );
});

test("login mode forces a headed browser", () => {
  const config = loadConfig(
    {
      SHAREPOINT_SITE_URL: "https://example.sharepoint.com/sites/genai",
      SHAREPOINT_PROFILE_DIR: path.join(path.sep, "tmp", "sp-mcp-profile"),
      SHAREPOINT_HEADLESS: "true",
    },
    { forceHeaded: true },
  );

  assert.equal(config.headless, false);
});

test("rejects normal Edge profile directories", () => {
  assert.throws(
    () =>
      assertSafeProfileDirectory(
        path.join(path.sep, "Users", "person", "AppData", "Local", "Microsoft", "Edge", "User Data", "Default"),
      ),
    /normal Edge profile/u,
  );
});
