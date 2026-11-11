import { EdgeSession } from "./browser/edge-session.js";
import { SharedEdgeManager } from "./browser/shared-edge.js";
import { loadConfig } from "./config.js";
import { StderrJsonLogger } from "./logger.js";

async function main(): Promise<void> {
  const config = loadConfig(process.env, { forceHeaded: true });
  const logger = new StderrJsonLogger();
  // Login needs the profile exclusively in a headed window, so stop the
  // shared Edge first. It restarts on demand at the next MCP tool call and
  // reuses the session this login refreshes.
  const stoppedSharedEdge = await new SharedEdgeManager(config, logger).stop();
  if (stoppedSharedEdge) {
    process.stderr.write(
      "Stopped the shared Edge process so the profile can be used for login.\n",
    );
  }
  const edgeSession = new EdgeSession(config, logger, "exclusive");
  const context = await edgeSession.getContext();
  const pages = context.pages();
  for (const [index, siteUrl] of config.siteUrls.entries()) {
    const page = index === 0
      ? (pages[0] ?? (await context.newPage()))
      : await context.newPage();
    try {
      await page.goto(siteUrl, {
        waitUntil: "domcontentloaded",
        timeout: config.requestTimeoutMs,
      });
    } catch (error) {
      logger.warn("login_navigation_incomplete", {
        siteUrl,
        errorType: error instanceof Error ? error.name : typeof error,
      });
    }
  }

  process.stderr.write(
    [
      "Microsoft Edge was opened with the dedicated SharePoint MCP profile.",
      `Complete SSO and MFA, confirm that all ${config.siteUrls.length} configured SharePoint sites open, then close the Edge window.`,
      "Do not use this profile for normal browsing.",
      "",
    ].join("\n"),
  );

  await new Promise<void>((resolve) => {
    context.once("close", () => resolve());
    process.once("SIGINT", () => {
      void edgeSession.close().finally(resolve);
    });
    process.once("SIGTERM", () => {
      void edgeSession.close().finally(resolve);
    });
  });
}

main().catch((error: unknown) => {
  const logger = new StderrJsonLogger();
  logger.error("login_command_failed", {
    errorType: error instanceof Error ? error.name : typeof error,
    message: error instanceof Error ? error.message : "Unknown login error",
  });
  process.exit(1);
});
