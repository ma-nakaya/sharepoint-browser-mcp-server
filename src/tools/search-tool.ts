import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

import type { SharePointReadService } from "../sharepoint/read-service.js";
import {
  DEFAULT_SEARCH_RESULTS,
  MAX_SEARCH_QUERY_CHARACTERS,
  MAX_SEARCH_RESULTS,
} from "../sharepoint/search.js";

export function registerSearchTool(
  server: McpServer,
  service: SharePointReadService,
): void {
  server.registerTool(
    "sharepoint_search",
    {
      title: "Search the configured SharePoint site",
      description: [
        "Searches content under the configured SharePoint site only.",
        "Returns titles, SharePoint URLs, content types, modified times, and short summaries.",
        "This tool is read-only and never returns cookies, tokens, or authorization headers.",
      ].join(" "),
      inputSchema: {
        query: z
          .string()
          .trim()
          .min(1)
          .max(MAX_SEARCH_QUERY_CHARACTERS)
          .describe("Keywords or a phrase to search for in the configured SharePoint site."),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(MAX_SEARCH_RESULTS)
          .optional()
          .describe(`Maximum results to return. Defaults to ${DEFAULT_SEARCH_RESULTS}.`),
      },
      outputSchema: {
        query: z.string(),
        siteUrl: z.string().url(),
        totalRows: z.number().int().nonnegative(),
        returnedRows: z.number().int().nonnegative(),
        method: z.enum(["api-request", "browser-fetch"]),
        results: z.array(
          z.object({
            title: z.string(),
            url: z.string().url(),
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
    async ({ query, maxResults }) => {
      const result = await service.search(
        query,
        maxResults ?? DEFAULT_SEARCH_RESULTS,
      );
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result, results: [...result.results] },
      };
    },
  );
}
