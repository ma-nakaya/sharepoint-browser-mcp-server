import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import {
  DEFAULT_FOLDER_RESULTS,
  MAX_DOWNLOAD_BYTES,
  MAX_FOLDER_RESULTS,
} from "../sharepoint/file-content.js";
import type { SharePointFileService } from "../sharepoint/file-service.js";

const requestMethodSchema = z.enum(["api-request", "browser-fetch"]);

export function registerFileTools(
  server: McpServer,
  service: SharePointFileService,
): void {
  registerDocumentLibrariesTool(server, service);
  registerFolderTool(server, service);
  registerDownloadTool(server, service);
}

function registerDocumentLibrariesTool(
  server: McpServer,
  service: SharePointFileService,
): void {
  server.registerTool(
    "sharepoint_list_document_libraries",
    {
      title: "List SharePoint document libraries",
      description: [
        "Lists visible document libraries under the configured SharePoint site.",
        "Use a returned root folder URL with sharepoint_list_folder.",
        "This tool is read-only.",
      ].join(" "),
      inputSchema: {},
      outputSchema: {
        siteUrl: z.string().url(),
        returnedLibraries: z.number().int().nonnegative(),
        method: requestMethodSchema,
        libraries: z.array(
          z.object({
            title: z.string(),
            url: z.string().url(),
            serverRelativeUrl: z.string(),
            itemCount: z.number().int().nonnegative(),
          }),
        ),
      },
      annotations: readOnlyAnnotations(),
    },
    async () => {
      const result = await service.listDocumentLibraries();
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: {
          ...result,
          libraries: [...result.libraries],
        },
      };
    },
  );
}

function registerFolderTool(
  server: McpServer,
  service: SharePointFileService,
): void {
  server.registerTool(
    "sharepoint_list_folder",
    {
      title: "List a SharePoint folder",
      description: [
        "Lists direct child folders and files for a folder under the configured SharePoint site.",
        "Use an absolute or server-relative folder URL returned by this server.",
        "The downloadable flag shows whether sharepoint_download_file accepts each file.",
      ].join(" "),
      inputSchema: {
        folderUrl: z
          .string()
          .trim()
          .min(1)
          .max(2_048)
          .describe("Absolute or server-relative URL of a folder under the configured site."),
        maxResults: z
          .number()
          .int()
          .min(1)
          .max(MAX_FOLDER_RESULTS)
          .optional()
          .describe(
            `Maximum folders and maximum files to return. Defaults to ${DEFAULT_FOLDER_RESULTS}.`,
          ),
      },
      outputSchema: {
        folderUrl: z.string().url(),
        serverRelativeUrl: z.string(),
        returnedFolders: z.number().int().nonnegative(),
        returnedFiles: z.number().int().nonnegative(),
        limitReached: z.boolean(),
        method: z.enum(["api-request", "browser-fetch", "mixed"]),
        folders: z.array(
          z.object({
            name: z.string(),
            url: z.string().url(),
            serverRelativeUrl: z.string(),
            itemCount: z.number().int().nonnegative(),
          }),
        ),
        files: z.array(
          z.object({
            name: z.string(),
            url: z.string().url(),
            serverRelativeUrl: z.string(),
            extension: z.string(),
            sizeBytes: z.number().int().nonnegative(),
            modifiedTime: z.string().optional(),
            majorVersion: z.number().int().nonnegative().optional(),
            downloadable: z.boolean(),
          }),
        ),
      },
      annotations: readOnlyAnnotations(),
    },
    async ({ folderUrl, maxResults }) => {
      const result = await service.listFolder(
        folderUrl,
        maxResults ?? DEFAULT_FOLDER_RESULTS,
      );
      return {
        content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        structuredContent: {
          ...result,
          folders: [...result.folders],
          files: [...result.files],
        },
      };
    },
  );
}

function registerDownloadTool(
  server: McpServer,
  service: SharePointFileService,
): void {
  server.registerTool(
    "sharepoint_download_file",
    {
      title: "Download a SharePoint file",
      description: [
        "Downloads an allowed document, text, or image file from the configured SharePoint site.",
        `Files larger than ${MAX_DOWNLOAD_BYTES} bytes and executable or active-content formats are rejected.`,
        "Returns the bytes as an embedded MCP resource and metadata separately.",
      ].join(" "),
      inputSchema: {
        fileUrl: z
          .string()
          .trim()
          .min(1)
          .max(2_048)
          .describe(
            "Absolute or server-relative URL of a downloadable file returned by sharepoint_list_folder.",
          ),
      },
      outputSchema: {
        name: z.string(),
        url: z.string().url(),
        serverRelativeUrl: z.string(),
        sizeBytes: z.number().int().nonnegative(),
        mimeType: z.string(),
        sha256: z.string().regex(/^[0-9a-f]{64}$/u),
        method: requestMethodSchema,
      },
      annotations: readOnlyAnnotations(),
    },
    async ({ fileUrl }) => {
      const result = await service.downloadFile(fileUrl);
      const metadata = {
        name: result.name,
        url: result.url,
        serverRelativeUrl: result.serverRelativeUrl,
        sizeBytes: result.sizeBytes,
        mimeType: result.mimeType,
        sha256: result.sha256,
        method: result.method,
      };
      return {
        content: [
          { type: "text" as const, text: JSON.stringify(metadata, null, 2) },
          {
            type: "resource" as const,
            resource: {
              uri: result.url,
              mimeType: result.mimeType,
              blob: Buffer.from(result.data).toString("base64"),
            },
          },
        ],
        structuredContent: metadata,
      };
    },
  );
}

function readOnlyAnnotations(): {
  readonly readOnlyHint: true;
  readonly destructiveHint: false;
  readonly idempotentHint: true;
  readonly openWorldHint: false;
} {
  return {
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  };
}
