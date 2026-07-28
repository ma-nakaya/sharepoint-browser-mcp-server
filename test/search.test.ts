import assert from "node:assert/strict";
import test from "node:test";

import {
  buildSearchApiPath,
  parseSearchResponse,
} from "../src/sharepoint/search.js";
import { buildSharePointApiUrl } from "../src/sharepoint/url-guard.js";

const SITE_URL = "https://example.sharepoint.com/sites/genai";

test("builds a site-scoped SharePoint search query", () => {
  const path = buildSearchApiPath(SITE_URL, "quarterly plan", 5);
  const url = buildSharePointApiUrl(SITE_URL, path);

  assert.equal(url.pathname, "/sites/genai/_api/search/query");
  assert.equal(url.searchParams.get("querytext"), "'quarterly plan'");
  assert.equal(
    url.searchParams.get("querytemplate"),
    `'({searchterms}) AND Path:"${SITE_URL}/*"'`,
  );
  assert.equal(url.searchParams.get("rowlimit"), "5");
  assert.equal(url.searchParams.get("rowsperpage"), "5");
  assert.equal(url.searchParams.get("startrow"), "0");
  assert.equal(url.searchParams.get("trimduplicates"), "true");
  assert.equal(url.searchParams.get("enablestemming"), "true");
  assert.match(url.searchParams.get("selectproperties") ?? "", /Title,Path/u);
});

test("builds a filtered, paged, and sorted site search query", () => {
  const path = buildSearchApiPath(SITE_URL, "annual report", 20, {
    startRow: 40,
    scope: "documents",
    folderUrl: "/sites/genai/Shared Documents/Finance",
    fileExtensions: [".PDF", "docx", "pdf"],
    modifiedAfter: "2025-04-01",
    modifiedBefore: "2026-03-31",
    sort: "modified-desc",
  });
  const url = buildSharePointApiUrl(SITE_URL, path);

  assert.equal(url.searchParams.get("startrow"), "40");
  assert.equal(url.searchParams.get("sortlist"), "'LastModifiedTime:descending'");
  assert.equal(
    url.searchParams.get("querytemplate"),
    [
      "'({searchterms})",
      'Path:\"https://example.sharepoint.com/sites/genai/Shared%20Documents/Finance/*\"',
      "IsDocument:True",
      "NOT FileType:aspx",
      "(FileType:pdf OR FileType:docx)",
      "LastModifiedTime>=2025-04-01",
      "LastModifiedTime<=2026-03-31'",
    ].join(" AND "),
  );
});

test("builds a page-only search query", () => {
  const path = buildSearchApiPath(SITE_URL, "news", 10, { scope: "pages" });
  const url = buildSharePointApiUrl(SITE_URL, path);

  assert.equal(
    url.searchParams.get("querytemplate"),
    `'({searchterms}) AND Path:"${SITE_URL}/*" AND FileType:aspx'`,
  );
});

test("rejects unsafe or contradictory search filters", () => {
  assert.throws(
    () =>
      buildSearchApiPath(SITE_URL, "policy", 10, {
        folderUrl: "https://example.sharepoint.com/sites/other/Documents",
      }),
    /configured SharePoint site/u,
  );
  assert.throws(
    () =>
      buildSearchApiPath(SITE_URL, "policy", 10, {
        scope: "pages",
        fileExtensions: ["pdf"],
      }),
    /cannot be combined/u,
  );
  assert.throws(
    () =>
      buildSearchApiPath(SITE_URL, "policy", 10, {
        modifiedAfter: "2026-02-30",
      }),
    /valid calendar date/u,
  );
  assert.throws(
    () =>
      buildSearchApiPath(SITE_URL, "policy", 10, {
        modifiedAfter: "2026-07-01",
        modifiedBefore: "2026-06-01",
      }),
    /must not be later/u,
  );
  assert.throws(
    () =>
      buildSearchApiPath(SITE_URL, "policy", 10, {
        fileExtensions: ["pdf OR Path:*"],
      }),
    /letters and numbers/u,
  );
});

test("escapes apostrophes in search query parameters", () => {
  const path = buildSearchApiPath(SITE_URL, "manager's guide", 10);
  const url = buildSharePointApiUrl(SITE_URL, path);

  assert.equal(url.searchParams.get("querytext"), "'manager''s guide'");
});

test("parses search rows and drops results outside the configured site", () => {
  const body = JSON.stringify({
    PrimaryQueryResult: {
      RelevantResults: {
        TotalRows: 2,
        Table: {
          Rows: {
            results: [
              {
                Cells: {
                  results: [
                    { Key: "Title", Value: "Quarterly plan" },
                    {
                      Key: "Path",
                      Value:
                        "https://example.sharepoint.com/sites/genai/SitePages/Plan.aspx",
                    },
                    { Key: "FileExtension", Value: "aspx" },
                    {
                      Key: "HitHighlightedSummary",
                      Value: "<c0>Quarterly</c0> <strong>plan</strong>",
                    },
                  ],
                },
              },
              {
                Cells: {
                  results: [
                    { Key: "Title", Value: "Other site" },
                    {
                      Key: "Path",
                      Value:
                        "https://example.sharepoint.com/sites/other/SitePages/Plan.aspx",
                    },
                  ],
                },
              },
            ],
          },
        },
      },
    },
  });

  const result = parseSearchResponse(body, SITE_URL, "quarterly", "api-request");

  assert.equal(result.totalRows, 2);
  assert.equal(result.returnedRows, 1);
  assert.equal(result.results[0]?.title, "Quarterly plan");
  assert.equal(result.results[0]?.kind, "page");
  assert.equal(result.results[0]?.fileExtension, "aspx");
  assert.equal(result.results[0]?.summary, "Quarterly plan");
});

test("parses verbose SharePoint search responses", () => {
  const body = JSON.stringify({
    d: {
      query: {
        PrimaryQueryResult: {
          RelevantResults: {
            TotalRows: "0",
            Table: { Rows: { results: [] } },
          },
        },
      },
    },
  });

  const result = parseSearchResponse(body, SITE_URL, "nothing", "browser-fetch");

  assert.equal(result.totalRows, 0);
  assert.equal(result.returnedRows, 0);
  assert.equal(result.method, "browser-fetch");
});

test("parses minimal-metadata responses with direct row and cell arrays", () => {
  const body = JSON.stringify({
    PrimaryQueryResult: {
      RelevantResults: {
        TotalRows: 1,
        Table: {
          Rows: [
            {
              Cells: [
                { Key: "Title", Value: "Home" },
                {
                  Key: "Path",
                  Value:
                    "https://example.sharepoint.com/sites/genai/SitePages/Home.aspx",
                },
              ],
            },
          ],
        },
      },
    },
  });

  const result = parseSearchResponse(body, SITE_URL, "home", "api-request");

  assert.equal(result.totalRows, 1);
  assert.equal(result.returnedRows, 1);
  assert.equal(result.results[0]?.title, "Home");
  assert.equal(result.results[0]?.kind, "other");
});

test("returns hierarchy metadata, result kinds, and a paging cursor", () => {
  const body = JSON.stringify({
    PrimaryQueryResult: {
      RelevantResults: {
        TotalRows: 45,
        Table: {
          Rows: [
            {
              Cells: [
                { Key: "Title", Value: "Annual report" },
                {
                  Key: "Path",
                  Value:
                    "https://example.sharepoint.com/sites/genai/Shared%20Documents/Annual.pdf",
                },
                {
                  Key: "ParentLink",
                  Value:
                    "https://example.sharepoint.com/sites/genai/Shared%20Documents",
                },
                { Key: "FileExtension", Value: "PDF" },
                { Key: "Author", Value: "A. Example" },
                { Key: "Size", Value: "2048" },
                { Key: "Rank", Value: "812.5" },
              ],
            },
          ],
        },
      },
    },
  });

  const result = parseSearchResponse(
    body,
    SITE_URL,
    "annual",
    "api-request",
    20,
    {
      startRow: 20,
      scope: "documents",
      fileExtensions: ["pdf"],
    },
  );

  assert.equal(result.startRow, 20);
  assert.equal(result.hasMore, true);
  assert.equal(result.nextStartRow, 40);
  assert.deepEqual(result.fileExtensions, ["pdf"]);
  assert.equal(result.results[0]?.kind, "document");
  assert.equal(
    result.results[0]?.parentUrl,
    "https://example.sharepoint.com/sites/genai/Shared%20Documents",
  );
  assert.equal(result.results[0]?.author, "A. Example");
  assert.equal(result.results[0]?.sizeBytes, 2048);
  assert.equal(result.results[0]?.rank, 812.5);
});
