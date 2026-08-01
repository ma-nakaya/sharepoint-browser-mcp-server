import assert from "node:assert/strict";
import test from "node:test";

import type { AuthStatusService } from "../src/sharepoint/auth-status-service.js";
import type { SharePointFileService } from "../src/sharepoint/file-service.js";
import {
  MultiSiteAuthStatusService,
  MultiSiteSharePointFileService,
  MultiSiteSharePointReadService,
  selectConfiguredSiteUrl,
} from "../src/sharepoint/multi-site-service.js";
import type { SharePointReadService } from "../src/sharepoint/read-service.js";
import type { SharePointSearchResult } from "../src/sharepoint/search.js";

const PRIMARY = "https://example.sharepoint.com/teams/ymsl";
const SOFTENG = "https://example.sharepoint.com/teams/ymsl_softeng";

test("routes absolute and server-relative resources to an exact configured site", () => {
  assert.equal(
    selectConfiguredSiteUrl(
      [PRIMARY, SOFTENG],
      `${SOFTENG}/SitePages/Home.aspx`,
    ),
    SOFTENG,
  );
  assert.equal(
    selectConfiguredSiteUrl(
      [PRIMARY, SOFTENG],
      "/teams/ymsl/Shared%20Documents/Policy.docx",
    ),
    PRIMARY,
  );
  assert.throws(
    () =>
      selectConfiguredSiteUrl(
        [PRIMARY, SOFTENG],
        "https://example.sharepoint.com/teams/other/SitePages/Home.aspx",
      ),
    /configured sites/u,
  );
});

test("combines ranked search results from all configured sites", async () => {
  const service = new MultiSiteSharePointReadService([
    { siteUrl: PRIMARY, service: fakeReadService(PRIMARY, "Primary", 10) },
    { siteUrl: SOFTENG, service: fakeReadService(SOFTENG, "Softeng", 20) },
  ]);

  const result = await service.search("design", 1);

  assert.deepEqual(result.siteUrls, [PRIMARY, SOFTENG]);
  assert.deepEqual(result.scopeUrls, [PRIMARY, SOFTENG]);
  assert.equal(result.totalRows, 2);
  assert.equal(result.returnedRows, 1);
  assert.equal(result.results[0]?.title, "Softeng");
  assert.equal(result.hasMore, true);
  await assert.rejects(
    () => service.search("design", 1, { startRow: 1 }),
    /requires siteUrl or folderUrl/u,
  );
});

test("routes a site-selected search without querying other sites", async () => {
  let primaryCalls = 0;
  let softengCalls = 0;
  const service = new MultiSiteSharePointReadService([
    {
      siteUrl: PRIMARY,
      service: fakeReadService(PRIMARY, "Primary", 10, () => {
        primaryCalls += 1;
      }),
    },
    {
      siteUrl: SOFTENG,
      service: fakeReadService(SOFTENG, "Softeng", 20, () => {
        softengCalls += 1;
      }),
    },
  ]);

  const result = await service.search("design", 10, { siteUrl: SOFTENG });

  assert.equal(primaryCalls, 0);
  assert.equal(softengCalls, 1);
  assert.deepEqual(result.siteUrls, [SOFTENG]);
});

test("aggregates authentication and document libraries across sites", async () => {
  const authService = new MultiSiteAuthStatusService([
    { siteUrl: PRIMARY, service: fakeAuthService(PRIMARY) },
    { siteUrl: SOFTENG, service: fakeAuthService(SOFTENG) },
  ]);
  const fileService = new MultiSiteSharePointFileService([
    { siteUrl: PRIMARY, service: fakeFileService(PRIMARY, "Documents") },
    { siteUrl: SOFTENG, service: fakeFileService(SOFTENG, "Engineering") },
  ]);

  const auth = await authService.getStatus();
  const libraries = await fileService.listDocumentLibraries();

  assert.equal(auth.authenticated, true);
  assert.deepEqual(auth.siteUrls, [PRIMARY, SOFTENG]);
  assert.equal(auth.sites.length, 2);
  assert.equal(libraries.returnedLibraries, 2);
  assert.deepEqual(libraries.siteUrls, [PRIMARY, SOFTENG]);
});

function fakeReadService(
  siteUrl: string,
  title: string,
  rank: number,
  onSearch: () => void = () => undefined,
): Pick<SharePointReadService, "search" | "getPage" | "getListItem"> {
  return {
    search: async (query, maxResults = 10, options = {}) => {
      onSearch();
      return searchResult(siteUrl, title, rank, query, maxResults, options.startRow ?? 0);
    },
    getPage: async () => {
      throw new Error("not used");
    },
    getListItem: async () => {
      throw new Error("not used");
    },
  };
}

function searchResult(
  siteUrl: string,
  title: string,
  rank: number,
  query: string,
  maxResults: number,
  startRow: number,
): SharePointSearchResult {
  return {
    query,
    siteUrl,
    totalRows: 1,
    returnedRows: 1,
    startRow,
    hasMore: false,
    scope: "all",
    scopeUrl: siteUrl,
    fileExtensions: [],
    sort: "relevance",
    method: "api-request",
    results: maxResults > 0
      ? [
          {
            title,
            url: `${siteUrl}/SitePages/Home.aspx`,
            kind: "page" as const,
            fileExtension: "aspx",
            rank,
          },
        ]
      : [],
  };
}

function fakeAuthService(
  siteUrl: string,
): Pick<AuthStatusService, "getStatus"> {
  return {
    getStatus: async () => ({
      authenticated: true,
      state: "AUTHENTICATED",
      siteUrl,
      method: "api-request",
      message: "Authenticated",
    }),
  };
}

function fakeFileService(
  siteUrl: string,
  title: string,
): Pick<
  SharePointFileService,
  "listDocumentLibraries" | "listFolder" | "downloadFile"
> {
  return {
    listDocumentLibraries: async () => ({
      siteUrl,
      returnedLibraries: 1,
      method: "api-request",
      libraries: [
        {
          title,
          url: `${siteUrl}/Documents`,
          serverRelativeUrl: `${new URL(siteUrl).pathname}/Documents`,
          itemCount: 1,
        },
      ],
    }),
    listFolder: async () => {
      throw new Error("not used");
    },
    downloadFile: async () => {
      throw new Error("not used");
    },
  };
}
