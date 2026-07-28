import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import { MAX_EXTRACTED_TEXT_CHARACTERS } from "../sharepoint/document-extractor.js";
import type { SharePointDocumentService } from "../sharepoint/document-service.js";

export function registerDocumentTool(
  server: McpServer,
  service: SharePointDocumentService,
): void {
  server.registerTool(
    "sharepoint_extract_document_text",
    {
      title: "Extract text from a SharePoint document",
      description: [
        "Downloads and extracts plain text from a PDF, DOCX, XLSX, or PPTX file under the configured SharePoint site.",
        `Extracted text is limited to ${MAX_EXTRACTED_TEXT_CHARACTERS} characters.`,
        "Document actions are not executed, Office archives are expanded with strict limits, and no file is saved locally.",
      ].join(" "),
      inputSchema: {
        fileUrl: z
          .string()
          .trim()
          .min(1)
          .max(2_048)
          .describe(
            "Absolute or server-relative URL of a PDF, DOCX, XLSX, or PPTX file returned by this server.",
          ),
      },
      outputSchema: {
        name: z.string(),
        url: z.string().url(),
        serverRelativeUrl: z.string(),
        sizeBytes: z.number().int().nonnegative(),
        mimeType: z.string(),
        sha256: z.string().regex(/^[0-9a-f]{64}$/u),
        method: z.enum(["api-request", "browser-fetch"]),
        format: z.enum(["pdf", "docx", "xlsx", "pptx"]),
        unitType: z.enum(["pages", "parts", "sheets", "slides"]),
        unitCount: z.number().int().nonnegative(),
        text: z.string(),
        characters: z.number().int().nonnegative(),
        truncated: z.boolean(),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ fileUrl }) => {
      const result = await service.extractText(fileUrl);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result },
      };
    },
  );
}
