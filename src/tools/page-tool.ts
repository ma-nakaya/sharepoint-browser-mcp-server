import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

import type { SharePointReadService } from "../sharepoint/read-service.js";

export function registerPageTool(
  server: McpServer,
  service: SharePointReadService,
): void {
  server.registerTool(
    "sharepoint_get_page",
    {
      title: "Read a SharePoint page",
      description: [
        "Reads authored text from an .aspx page in the configured site's SitePages library.",
        "Accepts an absolute SharePoint URL or a server-relative URL.",
        "Returns plain text only; raw CanvasContent1 markup and web-part configuration are not exposed.",
      ].join(" "),
      inputSchema: {
        pageUrl: z
          .string()
          .trim()
          .min(1)
          .max(2_048)
          .describe(
            "Absolute or server-relative URL of an .aspx page under the configured SitePages library.",
          ),
      },
      outputSchema: {
        title: z.string(),
        url: z.string().url(),
        serverRelativeUrl: z.string(),
        modifiedTime: z.string().optional(),
        text: z.string(),
        truncated: z.boolean(),
        method: z.enum(["api-request", "browser-fetch"]),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ pageUrl }) => {
      const result = await service.getPage(pageUrl);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result },
      };
    },
  );
}
