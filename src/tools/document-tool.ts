import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import {
  DEFAULT_DOCUMENT_SEARCH_RESULTS,
  MAX_DOCUMENT_NODE_IDS,
  MAX_DOCUMENT_SEARCH_QUERY_CHARACTERS,
  MAX_DOCUMENT_SEARCH_RESULTS,
  MAX_EXTRACTED_TEXT_CHARACTERS,
} from "../sharepoint/document-extractor.js";
import { MAX_DOCUMENT_SOURCE_BYTES } from "../sharepoint/file-content.js";
import type { SharePointDocumentService } from "../sharepoint/document-service.js";

const fileUrlSchema = z
  .string()
  .trim()
  .min(1)
  .max(2_048)
  .describe(
    "Absolute or server-relative URL of a PDF, DOCX, XLSX, or PPTX file returned by this server.",
  );

const expectedSha256Schema = z
  .string()
  .regex(/^[0-9a-fA-F]{64}$/u)
  .optional()
  .describe(
    "Optional SHA-256 from a prior outline or search result. The call fails if the SharePoint file changed.",
  );

const documentMetadataShape = {
  name: z.string(),
  url: z.string().url(),
  serverRelativeUrl: z.string(),
  sizeBytes: z.number().int().nonnegative(),
  mimeType: z.string(),
  sha256: z.string().regex(/^[0-9a-f]{64}$/u),
  method: z.enum(["api-request", "browser-fetch"]),
} as const;

const documentPositionShape = {
  nodeId: z.string(),
  title: z.string(),
  kind: z.enum(["page", "section", "part", "sheet", "slide"]),
  positionType: z.enum(["page", "paragraph", "part", "sheet", "slide"]),
  startIndex: z.number().int().nonnegative(),
  endIndex: z.number().int().nonnegative(),
  locator: z.string(),
} as const;

const outlineNodeSchema: z.ZodTypeAny = z.lazy(() =>
  z.object({
    ...documentPositionShape,
    characters: z.number().int().nonnegative(),
    truncated: z.boolean(),
    preview: z.string().optional(),
    children: z.array(outlineNodeSchema),
  }),
);

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
        `Document source files are limited to ${MAX_DOCUMENT_SOURCE_BYTES} bytes.`,
        `Extracted text is limited to ${MAX_EXTRACTED_TEXT_CHARACTERS} characters.`,
        "Document actions are not executed, Office archives are expanded with strict limits, and no file is saved locally.",
      ].join(" "),
      inputSchema: {
        fileUrl: fileUrlSchema,
      },
      outputSchema: {
        ...documentMetadataShape,
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

  server.registerTool(
    "sharepoint_get_document_outline",
    {
      title: "Get a structured SharePoint document outline",
      description: [
        "Builds a lightweight, deterministic outline for a PDF, DOCX, XLSX, or PPTX file.",
        `Document source files are limited to ${MAX_DOCUMENT_SOURCE_BYTES} bytes.`,
        "PDF pages, Word heading sections, Excel sheets, and PowerPoint slides receive stable node IDs and source locators.",
        "Returns short previews only; use sharepoint_get_document_nodes for selected full text.",
        "No external LLM is called and no index is saved locally.",
      ].join(" "),
      inputSchema: {
        fileUrl: fileUrlSchema,
      },
      outputSchema: {
        ...documentMetadataShape,
        format: z.enum(["pdf", "docx", "xlsx", "pptx"]),
        unitType: z.enum(["pages", "parts", "sheets", "slides"]),
        unitCount: z.number().int().nonnegative(),
        nodeCount: z.number().int().nonnegative(),
        truncated: z.boolean(),
        nodes: z.array(outlineNodeSchema),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ fileUrl }) => {
      const result = await service.getOutline(fileUrl);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result, nodes: [...result.nodes] },
      };
    },
  );

  server.registerTool(
    "sharepoint_search_document",
    {
      title: "Search inside a structured SharePoint document",
      description: [
        "Searches page, heading-section, sheet, or slide nodes inside one SharePoint document.",
        `Document source files are limited to ${MAX_DOCUMENT_SOURCE_BYTES} bytes.`,
        "Returns ranked node IDs, exact source locators, and compact snippets.",
        "Use sharepoint_get_document_nodes with selected node IDs to fetch focused text.",
        "Search is deterministic and local to the MCP process; no external LLM or persistent index is used.",
      ].join(" "),
      inputSchema: {
        fileUrl: fileUrlSchema,
        expectedSha256: expectedSha256Schema,
        query: z
          .string()
          .trim()
          .min(1)
          .max(MAX_DOCUMENT_SEARCH_QUERY_CHARACTERS)
          .describe("Keywords or phrase to find inside the selected document."),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(MAX_DOCUMENT_SEARCH_RESULTS)
          .optional()
          .describe(
            `Maximum matching nodes to return. Defaults to ${DEFAULT_DOCUMENT_SEARCH_RESULTS}.`,
          ),
      },
      outputSchema: {
        ...documentMetadataShape,
        format: z.enum(["pdf", "docx", "xlsx", "pptx"]),
        unitType: z.enum(["pages", "parts", "sheets", "slides"]),
        unitCount: z.number().int().nonnegative(),
        query: z.string(),
        matchedNodes: z.number().int().nonnegative(),
        returnedNodes: z.number().int().nonnegative(),
        truncated: z.boolean(),
        results: z.array(
          z.object({
            ...documentPositionShape,
            score: z.number().nonnegative(),
            snippet: z.string(),
          }),
        ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ fileUrl, expectedSha256, query, maxResults }) => {
      const result = await service.search(
        fileUrl,
        query,
        maxResults,
        expectedSha256,
      );
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result, results: [...result.results] },
      };
    },
  );

  server.registerTool(
    "sharepoint_get_document_nodes",
    {
      title: "Get selected SharePoint document nodes",
      description: [
        "Returns text only for selected server-generated document node IDs.",
        `Document source files are limited to ${MAX_DOCUMENT_SOURCE_BYTES} bytes.`,
        `Accepts up to ${MAX_DOCUMENT_NODE_IDS} nodes and limits combined text to ${MAX_EXTRACTED_TEXT_CHARACTERS} characters.`,
        "Call sharepoint_get_document_outline or sharepoint_search_document first to obtain valid node IDs.",
      ].join(" "),
      inputSchema: {
        fileUrl: fileUrlSchema,
        expectedSha256: expectedSha256Schema,
        nodeIds: z
          .array(
            z.string().regex(
              /^(?:page|section|part|sheet|slide)-\d{4}$/u,
            ),
          )
          .min(1)
          .max(MAX_DOCUMENT_NODE_IDS)
          .describe("Node IDs returned by this server."),
      },
      outputSchema: {
        ...documentMetadataShape,
        format: z.enum(["pdf", "docx", "xlsx", "pptx"]),
        unitType: z.enum(["pages", "parts", "sheets", "slides"]),
        unitCount: z.number().int().nonnegative(),
        requestedNodeIds: z.array(z.string()),
        returnedNodes: z.number().int().nonnegative(),
        characters: z.number().int().nonnegative(),
        truncated: z.boolean(),
        nodes: z.array(
          z.object({
            ...documentPositionShape,
            text: z.string(),
            characters: z.number().int().nonnegative(),
            truncated: z.boolean(),
          }),
        ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ fileUrl, expectedSha256, nodeIds }) => {
      const result = await service.getNodes(
        fileUrl,
        nodeIds,
        expectedSha256,
      );
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result, nodes: [...result.nodes] },
      };
    },
  );
}
