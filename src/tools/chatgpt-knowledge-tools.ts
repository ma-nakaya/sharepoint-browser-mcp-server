import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import type { ChatGptKnowledgeService } from "../sharepoint/chatgpt-knowledge-service.js";
import { MAX_SEARCH_QUERY_CHARACTERS } from "../sharepoint/search.js";

const searchItemSchema = z.object({
  id: z.string(),
  title: z.string(),
  url: z.string().url(),
});

const fetchMetadataValueSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
]);

export function registerChatGptKnowledgeTools(
  server: McpServer,
  service: ChatGptKnowledgeService,
): void {
  server.registerTool(
    "search",
    {
      title: "Search SharePoint knowledge",
      description: [
        "Use this first when the user wants to find knowledge across the configured SharePoint sites.",
        "For policies and regulations, search the exact title before trying document-only searches; current content may be a SharePoint list item rather than a file.",
        "Returns stable IDs, titles, and canonical URLs for citation.",
        "Call fetch with a returned ID to retrieve the full text.",
      ].join(" "),
      inputSchema: {
        query: z
          .string()
          .trim()
          .min(1)
          .max(MAX_SEARCH_QUERY_CHARACTERS)
          .describe("Keywords or a phrase to search for in SharePoint."),
      },
      outputSchema: {
        results: z.array(searchItemSchema),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ query }) => {
      const result = await service.search(query);
      const payload = { results: [...result.results] };
      return {
        structuredContent: payload,
        content: [{ type: "text" as const, text: JSON.stringify(payload) }],
      };
    },
  );

  server.registerTool(
    "fetch",
    {
      title: "Fetch SharePoint knowledge",
      description: [
        "Use this when full text is needed for a SharePoint result returned by search.",
        "Accepts a SharePoint SitePages page, Lists/DispForm.aspx list item, or supported document ID and returns canonical citation metadata.",
      ].join(" "),
      inputSchema: {
        id: z
          .string()
          .trim()
          .min(1)
          .max(2_048)
          .describe("The exact ID returned by the search tool."),
      },
      outputSchema: {
        id: z.string(),
        title: z.string(),
        text: z.string(),
        url: z.string().url(),
        metadata: z.record(fetchMetadataValueSchema),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ id }) => {
      try {
        const result = await service.fetch(id);
        const payload = {
          ...result,
          metadata: { ...result.metadata },
        };
        return {
          structuredContent: payload,
          content: [{ type: "text" as const, text: JSON.stringify(payload) }],
        };
      } catch (error) {
        return {
          isError: true,
          content: [
            {
              type: "text" as const,
              text: getSafeErrorMessage(error),
            },
          ],
        };
      }
    },
  );
}

function getSafeErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim()) {
    return error.message.trim().slice(0, 1_000);
  }
  return "SharePoint knowledge could not be fetched.";
}
