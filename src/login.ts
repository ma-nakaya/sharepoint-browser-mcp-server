import { EdgeSession } from "./browser/edge-session.js";
import { loadConfig } from "./config.js";
import { StderrJsonLogger } from "./logger.js";

async function main(): Promise<void> {
  const config = loadConfig(process.env, { forceHeaded: true });
  const logger = new StderrJsonLogger();
  const edgeSession = new EdgeSession(config, logger);
  const context = await edgeSession.getContext();
  const pages = context.pages();
  const page = pages[0] ?? (await context.newPage());

  try {
    await page.goto(config.siteUrl, {
      waitUntil: "domcontentloaded",
      timeout: config.requestTimeoutMs,
    });
  } catch (error) {
    logger.warn("login_navigation_incomplete", {
      errorType: error instanceof Error ? error.name : typeof error,
    });
  }

  process.stderr.write(
    [
      "Microsoft Edge was opened with the dedicated SharePoint MCP profile.",
      "Complete SSO and MFA, confirm that the configured SharePoint site opens, then close the Edge window.",
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
