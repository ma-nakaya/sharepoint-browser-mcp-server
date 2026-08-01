import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

import type { MultiSiteSharePointReadService } from "../sharepoint/multi-site-service.js";
import {
  DEFAULT_SEARCH_RESULTS,
  MAX_SEARCH_FILE_EXTENSIONS,
  MAX_SEARCH_QUERY_CHARACTERS,
  MAX_SEARCH_RESULTS,
  MAX_SEARCH_START_ROW,
} from "../sharepoint/search.js";

export function registerSearchTool(
  server: McpServer,
  service: Pick<MultiSiteSharePointReadService, "search">,
): void {
  server.registerTool(
    "sharepoint_search",
    {
      title: "Search the configured SharePoint sites",
      description: [
        "Searches pages and files across all configured SharePoint sites using SharePoint's own search index.",
        "Use siteUrl to select one configured site when deterministic paging is required.",
        "Supports folder, content-kind, file-extension, modified-date, sort, and paging controls.",
        "For policies and regulations, start with the exact title and scope=all because current content may be a Lists/DispForm.aspx item rather than a document file.",
        "Use the standard fetch tool for returned list-item URLs, sharepoint_get_page for SitePages URLs, and sharepoint_extract_document_text for PDF/DOCX/XLSX/PPTX URLs.",
        "This tool is read-only and never returns cookies, tokens, or authorization headers.",
      ].join(" "),
      inputSchema: {
        siteUrl: z
          .string()
          .url()
          .optional()
          .describe("Optional configured SharePoint site URL to search exclusively."),
        query: z
          .string()
          .trim()
          .min(1)
          .max(MAX_SEARCH_QUERY_CHARACTERS)
          .describe("Keywords or a phrase to search for across the configured SharePoint sites."),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(MAX_SEARCH_RESULTS)
          .optional()
          .describe(`Maximum results to return. Defaults to ${DEFAULT_SEARCH_RESULTS}.`),
        startRow: z
          .number()
          .int()
          .min(0)
          .max(MAX_SEARCH_START_ROW)
          .optional()
          .describe("Zero-based result offset for paging. Use nextStartRow from a previous result."),
        scope: z
          .enum(["all", "pages", "documents"])
          .optional()
          .describe("Limit results to SharePoint pages or documents. Defaults to all content."),
        folderUrl: z
          .string()
          .trim()
          .min(1)
          .max(2_048)
          .optional()
          .describe("Optional absolute or server-relative folder URL under the configured site."),
        fileExtensions: z
          .array(
            z
              .string()
              .trim()
              .min(1)
              .max(11)
              .regex(/^\.?[A-Za-z0-9]+$/u),
          )
          .max(MAX_SEARCH_FILE_EXTENSIONS)
          .optional()
          .describe("Optional extensions such as pdf, docx, xlsx, or pptx. Not valid with pages scope."),
        modifiedAfter: z
          .string()
          .regex(/^\d{4}-\d{2}-\d{2}$/u)
          .optional()
          .describe("Only return content modified on or after this YYYY-MM-DD date."),
        modifiedBefore: z
          .string()
          .regex(/^\d{4}-\d{2}-\d{2}$/u)
          .optional()
          .describe("Only return content modified on or before this YYYY-MM-DD date."),
        sort: z
          .enum(["relevance", "modified-desc"])
          .optional()
          .describe("Order by SharePoint relevance or newest modification time."),
      },
      outputSchema: {
        query: z.string(),
        siteUrl: z.string().url(),
        siteUrls: z.array(z.string().url()),
        totalRows: z.number().int().nonnegative(),
        returnedRows: z.number().int().nonnegative(),
        startRow: z.number().int().nonnegative(),
        hasMore: z.boolean(),
        nextStartRow: z.number().int().nonnegative().optional(),
        scope: z.enum(["all", "pages", "documents"]),
        scopeUrl: z.string().url(),
        scopeUrls: z.array(z.string().url()),
        fileExtensions: z.array(z.string()),
        modifiedAfter: z.string().optional(),
        modifiedBefore: z.string().optional(),
        sort: z.enum(["relevance", "modified-desc"]),
        method: z.enum(["api-request", "browser-fetch"]),
        results: z.array(
          z.object({
            title: z.string(),
            url: z.string().url(),
            kind: z.enum(["page", "document", "file", "other"]),
            parentUrl: z.string().url().optional(),
            author: z.string().optional(),
            sizeBytes: z.number().int().nonnegative().optional(),
            rank: z.number().optional(),
            fileExtension: z.string().optional(),
            contentClass: z.string().optional(),
            modifiedTime: z.string().optional(),
            summary: z.string().optional(),
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
    async ({
      siteUrl,
      query,
      maxResults,
      startRow,
      scope,
      folderUrl,
      fileExtensions,
      modifiedAfter,
      modifiedBefore,
      sort,
    }) => {
      const result = await service.search(
        query,
        maxResults ?? DEFAULT_SEARCH_RESULTS,
        {
          ...(startRow !== undefined ? { startRow } : {}),
          ...(siteUrl ? { siteUrl } : {}),
          ...(scope ? { scope } : {}),
          ...(folderUrl ? { folderUrl } : {}),
          ...(fileExtensions ? { fileExtensions } : {}),
          ...(modifiedAfter ? { modifiedAfter } : {}),
          ...(modifiedBefore ? { modifiedBefore } : {}),
          ...(sort ? { sort } : {}),
        },
      );
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result, results: [...result.results] },
      };
    },
  );
}
