import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

import { AUTH_STATES } from "../sharepoint/auth-status-service.js";
import type {
  MultiSiteAuthStatus,
  MultiSiteAuthStatusService,
} from "../sharepoint/multi-site-service.js";

export function registerAuthStatusTool(
  server: McpServer,
  service: Pick<MultiSiteAuthStatusService, "getStatus">,
): void {
  server.registerTool(
    "sharepoint_auth_status",
    {
      title: "SharePoint authentication status",
      description: [
        "Checks whether the dedicated Microsoft Edge profile is signed in",
        "and can read every configured SharePoint site.",
        "This tool never returns cookies, tokens, email addresses, or login names.",
      ].join(" "),
      inputSchema: {},
      outputSchema: {
        authenticated: z.boolean(),
        state: z.enum(AUTH_STATES),
        siteUrl: z.string().url(),
        siteUrls: z.array(z.string().url()),
        message: z.string(),
        method: z.enum(["api-request", "browser-fetch"]).optional(),
        user: z
          .object({
            id: z.number().int().nonnegative(),
            displayName: z.string(),
            isSiteAdmin: z.boolean(),
          })
          .optional(),
        sites: z.array(
          z.object({
            authenticated: z.boolean(),
            state: z.enum(AUTH_STATES),
            siteUrl: z.string().url(),
            message: z.string(),
            method: z.enum(["api-request", "browser-fetch"]).optional(),
            user: z
              .object({
                id: z.number().int().nonnegative(),
                displayName: z.string(),
                isSiteAdmin: z.boolean(),
              })
              .optional(),
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
    async () => {
      const result: MultiSiteAuthStatus = await service.getStatus();
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: { ...result },
      };
    },
  );
}
