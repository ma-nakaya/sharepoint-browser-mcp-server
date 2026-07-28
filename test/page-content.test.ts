import assert from "node:assert/strict";
import test from "node:test";

import {
  buildPageContentApiPath,
  canvasHtmlToPlainText,
  MAX_PAGE_TEXT_CHARACTERS,
  normalizeSharePointPageUrl,
  parsePageContentResponse,
} from "../src/sharepoint/page-content.js";
import { buildSharePointApiUrl } from "../src/sharepoint/url-guard.js";

const SITE_URL = "https://example.sharepoint.com/teams/genai";

test("normalizes absolute and server-relative SitePages URLs", () => {
  const absolute = normalizeSharePointPageUrl(
    SITE_URL,
    "https://example.sharepoint.com/teams/genai/SitePages/%E3%83%9B%E3%83%BC%E3%83%A0.aspx?source=nav#top",
  );
  const relative = normalizeSharePointPageUrl(
    SITE_URL,
    "/teams/genai/SitePages/News.aspx",
  );

  assert.equal(
    absolute.serverRelativeUrl,
    "/teams/genai/SitePages/ホーム.aspx",
  );
  assert.equal(
    absolute.pageUrl,
    "https://example.sharepoint.com/teams/genai/SitePages/%E3%83%9B%E3%83%BC%E3%83%A0.aspx",
  );
  assert.equal(
    relative.pageUrl,
    "https://example.sharepoint.com/teams/genai/SitePages/News.aspx",
  );
});

test("rejects pages outside the configured SitePages library", () => {
  assert.throws(
    () =>
      normalizeSharePointPageUrl(
        SITE_URL,
        "https://evil.example/teams/genai/SitePages/Home.aspx",
      ),
    /configured SharePoint origin/u,
  );
  assert.throws(
    () =>
      normalizeSharePointPageUrl(
        SITE_URL,
        "/teams/genai-other/SitePages/Home.aspx",
      ),
    /SitePages library/u,
  );
  assert.throws(
    () =>
      normalizeSharePointPageUrl(
        SITE_URL,
        "/teams/genai/Documents/Guide.docx",
      ),
    /SitePages library/u,
  );
});

test("builds a decoded ResourcePath REST endpoint", () => {
  const path = buildPageContentApiPath(
    "/teams/genai/SitePages/Manager's 100% # Guide.aspx",
  );
  const url = buildSharePointApiUrl(SITE_URL, path);

  assert.match(path, /GetFileByServerRelativePath/u);
  assert.match(path, /Manager%27%27s%20100%25%20%23%20Guide\.aspx/u);
  assert.match(path, /CanvasContent1/u);
  assert.equal(url.hash, "");
  assert.equal(url.pathname.includes("%23"), true);
});

test("extracts authored page text without exposing web-part attributes", () => {
  const html = [
    '<div data-sp-webpartdata="secret-configuration">',
    '<div data-sp-searchableplaintext="true">Mission &amp; Vision</div>',
    "</div>",
    "<script>hidden()</script>",
    "<p>First line<br>Second&nbsp;line</p>",
  ].join("");

  assert.equal(
    canvasHtmlToPlainText(html),
    ["Mission & Vision", "First line", "Second line"].join("\n"),
  );
  assert.doesNotMatch(canvasHtmlToPlainText(html), /secret|hidden/u);
});

test("parses page metadata and limits returned plain text", () => {
  const page = normalizeSharePointPageUrl(
    SITE_URL,
    "/teams/genai/SitePages/Long.aspx",
  );
  const body = JSON.stringify({
    Title: "Long page",
    FileRef: page.serverRelativeUrl,
    Modified: "2026-07-28T00:00:00Z",
    CanvasContent1: `<p>${"x".repeat(MAX_PAGE_TEXT_CHARACTERS + 10)}</p>`,
  });

  const result = parsePageContentResponse(body, page, "api-request");

  assert.equal(result.title, "Long page");
  assert.equal(result.text.length, MAX_PAGE_TEXT_CHARACTERS);
  assert.equal(result.truncated, true);
  assert.equal(result.method, "api-request");
});
