import assert from "node:assert/strict";
import test from "node:test";

import { KnowledgeRetrievalService } from "../src/sharepoint/knowledge-retrieval-service.js";

const SITE_URL = "https://example.sharepoint.com/sites/genai";
const PAGE_URL = `${SITE_URL}/SitePages/Home.aspx`;
const LIST_ITEM_URL = `${SITE_URL}/Lists/Rules/DispForm.aspx?ID=204`;
const DOCUMENT_URL = `${SITE_URL}/Shared%20Documents/Policy.docx`;

test("maps fetchable SharePoint results to the knowledge search contract", async () => {
  const readService = {
    search: async () => ({
      query: "policy",
      siteUrl: SITE_URL,
      totalRows: 5,
      returnedRows: 5,
      startRow: 0,
      hasMore: false,
      scope: "all" as const,
      scopeUrl: SITE_URL,
      fileExtensions: [],
      sort: "relevance" as const,
      method: "api-request" as const,
      results: [
        {
          title: "Home",
          url: PAGE_URL,
          kind: "page" as const,
          fileExtension: "aspx",
        },
        {
          title: "Policy",
          url: DOCUMENT_URL,
          kind: "document" as const,
          fileExtension: "DOCX",
        },
        {
          title: "Travel policy",
          url: LIST_ITEM_URL,
          kind: "page" as const,
          fileExtension: "aspx",
        },
        {
          title: "Image",
          url: `${SITE_URL}/Shared%20Documents/Image.png`,
          kind: "file" as const,
          fileExtension: "png",
        },
        {
          title: "Unknown",
          url: `${SITE_URL}/Lists/Items/1`,
          kind: "other" as const,
        },
      ],
    }),
    getPage: async () => {
      throw new Error("not used");
    },
    getListItem: async () => {
      throw new Error("not used");
    },
  };
  const documentService = {
    extractText: async () => {
      throw new Error("not used");
    },
  };
  const service = new KnowledgeRetrievalService(readService, documentService);

  const result = await service.search("policy");

  assert.deepEqual(result, {
    results: [
      { id: PAGE_URL, title: "Home", url: PAGE_URL },
      { id: DOCUMENT_URL, title: "Policy", url: DOCUMENT_URL },
      {
        id: LIST_ITEM_URL,
        title: "Travel policy",
        url: LIST_ITEM_URL,
      },
    ],
  });
});

test("fetches SharePoint page text with citation metadata", async () => {
  const readService = {
    search: async () => {
      throw new Error("not used");
    },
    getPage: async (id: string) => ({
      title: "Home",
      url: id,
      serverRelativeUrl: "/sites/genai/SitePages/Home.aspx",
      modifiedTime: "2026-07-29T01:23:45Z",
      text: "Welcome to the site.",
      truncated: false,
      method: "api-request" as const,
    }),
    getListItem: async () => {
      throw new Error("not used");
    },
  };
  const documentService = {
    extractText: async () => {
      throw new Error("not used");
    },
  };
  const service = new KnowledgeRetrievalService(readService, documentService);

  const result = await service.fetch(PAGE_URL);

  assert.equal(result.id, PAGE_URL);
  assert.equal(result.url, PAGE_URL);
  assert.equal(result.title, "Home");
  assert.equal(result.text, "Welcome to the site.");
  assert.deepEqual(result.metadata, {
    source: "sharepoint",
    kind: "page",
    serverRelativeUrl: "/sites/genai/SitePages/Home.aspx",
    modifiedTime: "2026-07-29T01:23:45Z",
    truncated: false,
  });
});

test("fetches SharePoint list item text with citation metadata", async () => {
  const readService = {
    search: async () => {
      throw new Error("not used");
    },
    getPage: async () => {
      throw new Error("not used");
    },
    getListItem: async (id: string) => ({
      title: "05 国内出張旅費規程",
      url: id,
      serverRelativeListUrl: "/sites/genai/Lists/Rules",
      itemId: 204,
      modifiedTime: "2026-07-29T01:23:45Z",
      text: "第1条 目的",
      fieldCount: 3,
      truncated: false,
      method: "api-request" as const,
    }),
  };
  const documentService = {
    extractText: async () => {
      throw new Error("not used");
    },
  };
  const service = new KnowledgeRetrievalService(readService, documentService);

  const result = await service.fetch(LIST_ITEM_URL);

  assert.equal(result.id, LIST_ITEM_URL);
  assert.equal(result.title, "05 国内出張旅費規程");
  assert.equal(result.text, "第1条 目的");
  assert.deepEqual(result.metadata, {
    source: "sharepoint",
    kind: "list-item",
    serverRelativeListUrl: "/sites/genai/Lists/Rules",
    itemId: 204,
    fieldCount: 3,
    modifiedTime: "2026-07-29T01:23:45Z",
    truncated: false,
  });
});

test("fetches supported SharePoint document text with citation metadata", async () => {
  const readService = {
    search: async () => {
      throw new Error("not used");
    },
    getPage: async () => {
      throw new Error("not used");
    },
    getListItem: async () => {
      throw new Error("not used");
    },
  };
  const documentService = {
    extractText: async (id: string) => ({
      name: "Policy.docx",
      url: id,
      serverRelativeUrl: "/sites/genai/Shared Documents/Policy.docx",
      sizeBytes: 1234,
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      sha256: "a".repeat(64),
      method: "browser-fetch" as const,
      format: "docx" as const,
      unitType: "parts" as const,
      unitCount: 1,
      text: "Company policy.",
      characters: 15,
      truncated: false,
    }),
  };
  const service = new KnowledgeRetrievalService(readService, documentService);

  const result = await service.fetch(DOCUMENT_URL);

  assert.equal(result.id, DOCUMENT_URL);
  assert.equal(result.title, "Policy.docx");
  assert.equal(result.text, "Company policy.");
  assert.deepEqual(result.metadata, {
    source: "sharepoint",
    kind: "document",
    format: "docx",
    unitType: "parts",
    unitCount: 1,
    sizeBytes: 1234,
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    sha256: "a".repeat(64),
    truncated: false,
  });
});

test("rejects unsupported knowledge item IDs before SharePoint access", async () => {
  let pageCalls = 0;
  let listItemCalls = 0;
  let documentCalls = 0;
  const service = new KnowledgeRetrievalService(
    {
      search: async () => {
        throw new Error("not used");
      },
      getPage: async () => {
        pageCalls += 1;
        throw new Error("unexpected");
      },
      getListItem: async () => {
        listItemCalls += 1;
        throw new Error("unexpected");
      },
    },
    {
      extractText: async () => {
        documentCalls += 1;
        throw new Error("unexpected");
      },
    },
  );

  await assert.rejects(
    () => service.fetch(`${SITE_URL}/Shared%20Documents/image.png`),
    /supported PDF/u,
  );
  assert.equal(pageCalls, 0);
  assert.equal(listItemCalls, 0);
  assert.equal(documentCalls, 0);
});
