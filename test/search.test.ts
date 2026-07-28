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
  assert.equal(url.searchParams.get("trimduplicates"), "true");
  assert.match(url.searchParams.get("selectproperties") ?? "", /Title,Path/u);
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
});
