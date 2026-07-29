import assert from "node:assert/strict";
import test from "node:test";

import { strToU8, zipSync } from "fflate";

import { SharePointDocumentService } from "../src/sharepoint/document-service.js";
import { MAX_DOCUMENT_SOURCE_BYTES } from "../src/sharepoint/file-content.js";
import type { SharePointFileService } from "../src/sharepoint/file-service.js";

test("combines SharePoint file metadata with extracted text", async () => {
  const data = zipSync({
    "word/document.xml": strToU8(
      '<w:document xmlns:w="urn:w"><w:body><w:p><w:t>Policy text</w:t></w:p></w:body></w:document>',
    ),
  });
  let requestedLimit: number | undefined;
  const fileService = {
    downloadFile: async (_fileUrl: string, maxBytes?: number) => {
      requestedLimit = maxBytes;
      return {
        name: "Policy.docx",
        url: "https://example.sharepoint.com/sites/genai/Documents/Policy.docx",
        serverRelativeUrl: "/sites/genai/Documents/Policy.docx",
        extension: ".docx",
        sizeBytes: data.byteLength,
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        sha256: "a".repeat(64),
        method: "api-request" as const,
        data,
      };
    },
  };
  const service = new SharePointDocumentService(
    fileService as unknown as SharePointFileService,
  );

  const result = await service.extractText(
    "/sites/genai/Documents/Policy.docx",
  );

  assert.equal(result.name, "Policy.docx");
  assert.equal(result.format, "docx");
  assert.match(result.text, /Policy text/u);
  assert.equal(result.characters, result.text.length);
  assert.equal(requestedLimit, MAX_DOCUMENT_SOURCE_BYTES);
});

test("combines file identity with outline, search, and selected nodes", async () => {
  const data = zipSync({
    "word/document.xml": strToU8(
      [
        '<w:document xmlns:w="urn:w"><w:body>',
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:t>Security</w:t></w:p>',
        "<w:p><w:t>Use least privilege</w:t></w:p>",
        "</w:body></w:document>",
      ].join(""),
    ),
  });
  const fileService = {
    downloadFile: async () => ({
      name: "Security.docx",
      url: "https://example.sharepoint.com/sites/genai/Documents/Security.docx",
      serverRelativeUrl: "/sites/genai/Documents/Security.docx",
      extension: ".docx",
      sizeBytes: data.byteLength,
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      sha256: "b".repeat(64),
      method: "browser-fetch" as const,
      data,
    }),
  };
  const service = new SharePointDocumentService(
    fileService as unknown as SharePointFileService,
  );

  const outline = await service.getOutline("/sites/genai/Documents/Security.docx");
  const search = await service.search(
    "/sites/genai/Documents/Security.docx",
    "least privilege",
  );
  const selected = await service.getNodes(
    "/sites/genai/Documents/Security.docx",
    ["section-0001"],
    "B".repeat(64),
  );

  assert.equal(outline.sha256, "b".repeat(64));
  assert.equal(outline.method, "browser-fetch");
  assert.equal(outline.nodes[0]?.title, "Security");
  assert.equal(search.results[0]?.nodeId, "section-0001");
  assert.match(selected.nodes[0]?.text ?? "", /least privilege/u);
  await assert.rejects(
    () =>
      service.getNodes(
        "/sites/genai/Documents/Security.docx",
        ["section-0001"],
        "c".repeat(64),
      ),
    /document changed/u,
  );
});
