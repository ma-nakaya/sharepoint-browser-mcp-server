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
        "Use this when the user wants to find pages or documents in the configured SharePoint site.",
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
        "Accepts only a SharePoint page or supported document ID and returns canonical citation metadata.",
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
      const result = await service.fetch(id);
      const payload = {
        ...result,
        metadata: { ...result.metadata },
      };
      return {
        structuredContent: payload,
        content: [{ type: "text" as const, text: JSON.stringify(payload) }],
      };
    },
  );
}
