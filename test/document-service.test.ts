import assert from "node:assert/strict";
import test from "node:test";

import { strToU8, zipSync } from "fflate";

import { SharePointDocumentService } from "../src/sharepoint/document-service.js";
import type { SharePointFileService } from "../src/sharepoint/file-service.js";

test("combines SharePoint file metadata with extracted text", async () => {
  const data = zipSync({
    "word/document.xml": strToU8(
      '<w:document xmlns:w="urn:w"><w:body><w:p><w:t>Policy text</w:t></w:p></w:body></w:document>',
    ),
  });
  const fileService = {
    downloadFile: async () => ({
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
    }),
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
});
