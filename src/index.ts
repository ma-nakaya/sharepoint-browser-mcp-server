import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { EdgeSession } from "./browser/edge-session.js";
import { loadConfig } from "./config.js";
import { StderrJsonLogger } from "./logger.js";
import { AuthStatusService } from "./sharepoint/auth-status-service.js";
import { PlaywrightSharePointTransport } from "./sharepoint/playwright-transport.js";
import { registerAuthStatusTool } from "./tools/auth-status-tool.js";

const SERVER_NAME = "sharepoint-browser-mcp-server";
const SERVER_VERSION = "0.1.0";

async function main(): Promise<void> {
  const config = loadConfig();
  const logger = new StderrJsonLogger();
  const edgeSession = new EdgeSession(config, logger);
  const sharePointTransport = new PlaywrightSharePointTransport(config, edgeSession, logger);
  const authStatusService = new AuthStatusService(config.siteUrl, sharePointTransport);

  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    {
      instructions: [
        "This server is read-only.",
        "Call sharepoint_auth_status before SharePoint operations.",
        "Never request or expose browser cookies, authorization headers, or tokens.",
      ].join(" "),
    },
  );

  registerAuthStatusTool(server, authStatusService);

  let shuttingDown = false;
  const shutdown = async (signal: string): Promise<void> => {
    if (shuttingDown) {
      return;
    }
    shuttingDown = true;
    logger.info("server_stopping", { signal });
    await sharePointTransport.close();
    process.exit(0);
  };

  process.once("SIGINT", () => void shutdown("SIGINT"));
  process.once("SIGTERM", () => void shutdown("SIGTERM"));

  process.on("uncaughtException", (error) => {
    logger.error("uncaught_exception", { errorType: error.name });
    void sharePointTransport.close().finally(() => process.exit(1));
  });

  process.on("unhandledRejection", (reason) => {
    logger.error("unhandled_rejection", {
      errorType: reason instanceof Error ? reason.name : typeof reason,
    });
    void sharePointTransport.close().finally(() => process.exit(1));
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  logger.info("server_started", {
    transport: "stdio",
    siteOrigin: config.siteOrigin,
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
