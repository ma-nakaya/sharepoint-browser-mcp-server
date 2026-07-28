import assert from "node:assert/strict";
import test from "node:test";

import type {
  SharePointResponse,
  SharePointTransport,
} from "../src/sharepoint/http.js";
import { SharePointReadService } from "../src/sharepoint/read-service.js";

const SITE_URL = "https://example.sharepoint.com/sites/genai";

class FakeTransport implements SharePointTransport {
  public primaryCalls = 0;
  public fallbackCalls = 0;
  public lastApiPath: string | undefined;

  constructor(
    private readonly primary: SharePointResponse | Error,
    private readonly fallback: SharePointResponse | Error,
  ) {}

  async get(apiPath: string): Promise<SharePointResponse> {
    this.primaryCalls += 1;
    this.lastApiPath = apiPath;
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
  bodyTruncated = false,
): SharePointResponse {
  return { status, body, contentType, method, bodyTruncated };
}

function emptySearchBody(): string {
  return JSON.stringify({
    PrimaryQueryResult: {
      RelevantResults: {
        TotalRows: 0,
        Table: { Rows: { results: [] } },
      },
    },
  });
}

test("returns search results from the primary API request", async () => {
  const transport = new FakeTransport(
    response(200, emptySearchBody()),
    new Error("fallback must not run"),
  );
  const service = new SharePointReadService(SITE_URL, transport);

  const result = await service.search("policy");

  assert.equal(result.returnedRows, 0);
  assert.equal(result.method, "api-request");
  assert.equal(transport.fallbackCalls, 0);
});

test("passes site-search filters and paging to SharePoint", async () => {
  const transport = new FakeTransport(
    response(200, emptySearchBody()),
    new Error("fallback must not run"),
  );
  const service = new SharePointReadService(SITE_URL, transport);

  const result = await service.search("policy", 15, {
    startRow: 30,
    scope: "documents",
    fileExtensions: ["pdf"],
    sort: "modified-desc",
  });
  const url = new URL(transport.lastApiPath ?? "", `${SITE_URL}/`);

  assert.equal(url.searchParams.get("rowlimit"), "15");
  assert.equal(url.searchParams.get("startrow"), "30");
  assert.match(url.searchParams.get("querytemplate") ?? "", /FileType:pdf/u);
  assert.equal(url.searchParams.get("sortlist"), "'LastModifiedTime:descending'");
  assert.equal(result.startRow, 30);
  assert.equal(result.scope, "documents");
});

test("uses browser fetch for a non-JSON primary response", async () => {
  const transport = new FakeTransport(
    response(200, "<html>login</html>", "text/html"),
    response(
      200,
      emptySearchBody(),
      "application/json",
      "browser-fetch",
    ),
  );
  const service = new SharePointReadService(SITE_URL, transport);

  const result = await service.search("policy");

  assert.equal(result.method, "browser-fetch");
  assert.equal(transport.fallbackCalls, 1);
});

test("does not retry a clear page not-found response", async () => {
  const transport = new FakeTransport(
    response(404),
    new Error("fallback must not run"),
  );
  const service = new SharePointReadService(SITE_URL, transport);

  await assert.rejects(
    () => service.getPage("/sites/genai/SitePages/Missing.aspx"),
    /page was not found/u,
  );
  assert.equal(transport.fallbackCalls, 0);
});

test("reads a normalized SharePoint list item through FieldValuesAsText", async () => {
  const transport = new FakeTransport(
    response(
      200,
      JSON.stringify({
        Title: "05 国内出張旅費規程",
        Body: "<p>第1条 目的</p>",
      }),
    ),
    new Error("fallback must not run"),
  );
  const service = new SharePointReadService(SITE_URL, transport);

  const result = await service.getListItem(
    "/sites/genai/Lists/Rules/DispForm.aspx?ID=204",
  );
  const url = new URL(transport.lastApiPath ?? "", `${SITE_URL}/`);

  assert.equal(result.title, "05 国内出張旅費規程");
  assert.equal(result.itemId, 204);
  assert.match(result.text, /第1条 目的/u);
  assert.equal(
    url.pathname,
    "/_api/web/GetList(@listUrl)/items(204)/FieldValuesAsText",
  );
  assert.equal(
    url.searchParams.get("@listUrl"),
    "'/sites/genai/Lists/Rules'",
  );
  assert.equal(transport.fallbackCalls, 0);
});

test("rejects truncated JSON responses without parsing the body", async () => {
  const transport = new FakeTransport(
    response(200, '{"partial":', "application/json", "api-request", true),
    new Error("fallback must not run"),
  );
  const service = new SharePointReadService(SITE_URL, transport);

  await assert.rejects(
    () => service.search("policy"),
    /safe response limit/u,
  );
  assert.equal(transport.fallbackCalls, 0);
});
