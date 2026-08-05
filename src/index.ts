import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { EdgeSession } from "./browser/edge-session.js";
import { loadConfig } from "./config.js";
import { StderrJsonLogger } from "./logger.js";
import { AuthStatusService } from "./sharepoint/auth-status-service.js";
import { KnowledgeRetrievalService } from "./sharepoint/knowledge-retrieval-service.js";
import { SharePointDocumentService } from "./sharepoint/document-service.js";
import { SharePointFileService } from "./sharepoint/file-service.js";
import {
  MultiSiteAuthStatusService,
  MultiSiteSharePointFileService,
  MultiSiteSharePointReadService,
} from "./sharepoint/multi-site-service.js";
import { PlaywrightSharePointTransport } from "./sharepoint/playwright-transport.js";
import { SharePointReadService } from "./sharepoint/read-service.js";
import { registerAuthStatusTool } from "./tools/auth-status-tool.js";
import { registerKnowledgeRetrievalTools } from "./tools/knowledge-retrieval-tools.js";
import { registerDocumentTool } from "./tools/document-tool.js";
import { registerFileTools } from "./tools/file-tools.js";
import { registerPageTool } from "./tools/page-tool.js";
import { registerSearchTool } from "./tools/search-tool.js";

const SERVER_NAME = "sharepoint-browser-mcp-server";
const SERVER_VERSION = "0.10.0";

async function main(): Promise<void> {
  const config = loadConfig();
  const logger = new StderrJsonLogger();
  const edgeSession = new EdgeSession(config, logger);
  const siteServices = config.siteUrls.map((siteUrl) => {
    const siteConfig = {
      ...config,
      siteUrl,
      siteOrigin: new URL(siteUrl).origin,
    };
    const transport = new PlaywrightSharePointTransport(
      siteConfig,
      edgeSession,
      logger,
    );
    return {
      siteUrl,
      authStatusService: new AuthStatusService(siteUrl, transport),
      readService: new SharePointReadService(siteUrl, transport),
      fileService: new SharePointFileService(siteUrl, transport),
    };
  });
  const authStatusService = new MultiSiteAuthStatusService(
    siteServices.map(({ siteUrl, authStatusService: service }) => ({
      siteUrl,
      service,
    })),
  );
  const readService = new MultiSiteSharePointReadService(
    siteServices.map(({ siteUrl, readService: service }) => ({
      siteUrl,
      service,
    })),
  );
  const fileService = new MultiSiteSharePointFileService(
    siteServices.map(({ siteUrl, fileService: service }) => ({
      siteUrl,
      service,
    })),
  );
  const documentService = new SharePointDocumentService(fileService);
  const knowledgeRetrievalService = new KnowledgeRetrievalService(
    readService,
    documentService,
  );

  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    {
      instructions: [
        "This server is read-only.",
        "Call sharepoint_auth_status before SharePoint operations.",
        "Use search and fetch across the configured SharePoint sites for portable read-only knowledge retrieval.",
        "For policies and regulations, search the exact title with the standard search tool first; fetch supports both SitePages and Lists/DispForm.aspx results.",
        "Use sharepoint_search to find content across configured sites before opening pages or extracting documents.",
        "For large PDF or Office files, inspect the document outline or search document nodes before fetching selected node text.",
        "Never request or expose browser cookies, authorization headers, or tokens.",
      ].join(" "),
    },
  );

  registerAuthStatusTool(server, authStatusService);
  registerKnowledgeRetrievalTools(server, knowledgeRetrievalService);
  registerSearchTool(server, readService);
  registerPageTool(server, readService);
  registerFileTools(server, fileService);
  registerDocumentTool(server, documentService);

  let shuttingDown = false;
  const shutdown = async (signal: string): Promise<void> => {
    if (shuttingDown) {
      return;
    }
    shuttingDown = true;
    logger.info("server_stopping", { signal });
    await edgeSession.close();
    process.exit(0);
  };

  process.once("SIGINT", () => void shutdown("SIGINT"));
  process.once("SIGTERM", () => void shutdown("SIGTERM"));

  process.on("uncaughtException", (error) => {
    logger.error("uncaught_exception", { errorType: error.name });
    void edgeSession.close().finally(() => process.exit(1));
  });

  process.on("unhandledRejection", (reason) => {
    logger.error("unhandled_rejection", {
      errorType: reason instanceof Error ? reason.name : typeof reason,
    });
    void edgeSession.close().finally(() => process.exit(1));
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  logger.info("server_started", {
    transport: "stdio",
    siteOrigin: config.siteOrigin,
    siteUrls: config.siteUrls,
    headless: config.headless,
  });
}

main().catch((error: unknown) => {
  const logger = new StderrJsonLogger();
  logger.error("server_start_failed", {
    errorType: error instanceof Error ? error.name : typeof error,
    message: error instanceof Error ? error.message : "Unknown startup error",
  });
  process.exit(1);
});
