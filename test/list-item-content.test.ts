import assert from "node:assert/strict";
import test from "node:test";

import {
  buildListItemContentApiPath,
  isSharePointListItemFormUrl,
  MAX_LIST_ITEM_TEXT_CHARACTERS,
  normalizeSharePointListItemUrl,
  parseListItemContentResponse,
} from "../src/sharepoint/list-item-content.js";
import { buildSharePointApiUrl } from "../src/sharepoint/url-guard.js";

const SITE_URL = "https://example.sharepoint.com/teams/genai";
const ITEM_URL =
  `${SITE_URL}/Lists/Rules/DispForm.aspx?ID=204`;

test("recognizes and normalizes SharePoint list item display-form URLs", () => {
  const absolute = normalizeSharePointListItemUrl(
    SITE_URL,
    `${ITEM_URL}&Source=%2Fteams%2Fgenai%2FLists%2FRules%2FAllItems.aspx#top`,
  );
  const relative = normalizeSharePointListItemUrl(
    SITE_URL,
    "/teams/genai/Lists/%E8%A6%8F%E7%A8%8B/DispForm.aspx?id=206",
  );

  assert.equal(absolute.itemUrl, ITEM_URL);
  assert.equal(absolute.serverRelativeListUrl, "/teams/genai/Lists/Rules");
  assert.equal(absolute.itemId, 204);
  assert.equal(
    relative.itemUrl,
    `${SITE_URL}/Lists/%E8%A6%8F%E7%A8%8B/DispForm.aspx?ID=206`,
  );
  assert.equal(relative.serverRelativeListUrl, "/teams/genai/Lists/規程");
  assert.equal(relative.itemId, 206);
  assert.equal(isSharePointListItemFormUrl(absolute.itemUrl), true);
  assert.equal(
    isSharePointListItemFormUrl(`${SITE_URL}/SitePages/Home.aspx`),
    false,
  );
});

test("rejects list item URLs outside the configured list boundary", () => {
  assert.throws(
    () =>
      normalizeSharePointListItemUrl(
        SITE_URL,
        "https://evil.example/teams/genai/Lists/Rules/DispForm.aspx?ID=1",
      ),
    /configured SharePoint origin/u,
  );
  assert.throws(
    () =>
      normalizeSharePointListItemUrl(
        SITE_URL,
        "/teams/other/Lists/Rules/DispForm.aspx?ID=1",
      ),
    /configured site's Lists path/u,
  );
  assert.throws(
    () =>
      normalizeSharePointListItemUrl(
        SITE_URL,
        "/teams/genai/SitePages/DispForm.aspx?ID=1",
      ),
    /configured site's Lists path/u,
  );
  assert.throws(
    () =>
      normalizeSharePointListItemUrl(
        SITE_URL,
        "/teams/genai/Lists/Rules/Forms/DispForm.aspx?ID=1",
      ),
    /directly under a SharePoint list/u,
  );
});

test("requires one positive integer list item ID", () => {
  for (const url of [
    "/teams/genai/Lists/Rules/DispForm.aspx",
    "/teams/genai/Lists/Rules/DispForm.aspx?ID=0",
    "/teams/genai/Lists/Rules/DispForm.aspx?ID=-1",
    "/teams/genai/Lists/Rules/DispForm.aspx?ID=1.5",
    "/teams/genai/Lists/Rules/DispForm.aspx?ID=1&id=2",
  ]) {
    assert.throws(
      () => normalizeSharePointListItemUrl(SITE_URL, url),
      /exactly one positive integer ID/u,
    );
  }
});

test("builds a list-scoped FieldValuesAsText endpoint", () => {
  const item = normalizeSharePointListItemUrl(
    SITE_URL,
    "/teams/genai/Lists/Manager's Rules/DispForm.aspx?ID=204",
  );
  const path = buildListItemContentApiPath(item);
  const url = buildSharePointApiUrl(SITE_URL, path);

  assert.equal(
    url.pathname,
    "/teams/genai/_api/web/GetList(@listUrl)/items(204)/FieldValuesAsText",
  );
  assert.equal(
    url.searchParams.get("@listUrl"),
    "'/teams/genai/Lists/Manager''s Rules'",
  );
});

test("parses authored list fields and drops system metadata", () => {
  const item = normalizeSharePointListItemUrl(SITE_URL, ITEM_URL);
  const body = JSON.stringify({
    Title: "05 国内出張旅費規程",
    Modified: "2026-07-28T00:00:00Z",
    OData_x005f_x5185_x005f_x5bb9_:
      "<p>第1条&nbsp;目的</p><p>第2条&nbsp;適用範囲</p>",
    Category_x0020_Name: "旅費;#国内",
    Editor: "Hidden User",
    ContentVersion: "0",
    ContentTypeId: "0x0100",
    OData__UIVersion: "512",
    OData_Created_x005f_x0020_x005f_x0020_Date: "2015/11/13 17:27",
    NoExecute: "0",
    SMTotalFileStreamSize: "0",
    ComplianceAssetId: "retention metadata",
    MetaInfo: "opaque metadata",
    SyncClientId: "sync metadata",
    ParentUniqueId: "{AB4E686E-37D8-423C-8F03-73B8198039AA}",
    TaxCatchAll: "opaque-taxonomy-data",
    e7ebc32b391a492a9472ebef8269e15c: "hidden taxonomy field",
    __metadata: { type: "SP.FieldStringValues" },
  });

  const result = parseListItemContentResponse(body, item, "api-request");

  assert.equal(result.title, "05 国内出張旅費規程");
  assert.equal(result.modifiedTime, "2026-07-28T00:00:00Z");
  assert.equal(result.itemId, 204);
  assert.equal(result.serverRelativeListUrl, "/teams/genai/Lists/Rules");
  assert.equal(result.fieldCount, 2);
  assert.equal(
    result.text,
    [
      "05 国内出張旅費規程",
      "内容:\n第1条 目的\n第2条 適用範囲",
      "Category Name: 旅費; 国内",
    ].join("\n\n"),
  );
  assert.doesNotMatch(
    result.text,
    /Hidden User|ContentVersion|ContentTypeId|UIVersion|Created|NoExecute|SMTotal|Compliance|MetaInfo|SyncClient|ParentUniqueId|TaxCatchAll|taxonomy/u,
  );
  assert.equal(result.truncated, false);
});

test("parses verbose expanded FieldValuesAsText and caps returned text", () => {
  const item = normalizeSharePointListItemUrl(SITE_URL, ITEM_URL);
  const body = JSON.stringify({
    d: {
      FieldValuesAsText: {
        Title: "Long regulation",
        RegulationBody: "x".repeat(MAX_LIST_ITEM_TEXT_CHARACTERS + 100),
      },
    },
  });

  const result = parseListItemContentResponse(body, item, "browser-fetch");

  assert.equal(result.text.length, MAX_LIST_ITEM_TEXT_CHARACTERS);
  assert.equal(result.truncated, true);
  assert.equal(result.method, "browser-fetch");
});
