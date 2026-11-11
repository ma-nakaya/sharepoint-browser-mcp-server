import { SharedEdgeManager } from "./browser/shared-edge.js";
import {
  DEFAULT_CDP_PORT,
  resolveProfileDir,
} from "./config.js";
import { StderrJsonLogger } from "./logger.js";

async function main(): Promise<void> {
  const logger = new StderrJsonLogger();
  const rawPort = process.env.SHAREPOINT_CDP_PORT?.trim();
  const parsedPort = rawPort === undefined ? Number.NaN : Number(rawPort);
  const manager = new SharedEdgeManager(
    {
      profileDir: resolveProfileDir(
        process.env.SHAREPOINT_PROFILE_DIR,
        process.env,
      ),
      cdpPort: Number.isSafeInteger(parsedPort) ? parsedPort : DEFAULT_CDP_PORT,
      headless: true,
      browserChannel: "msedge",
      edgeExecutablePath: undefined,
    },
    logger,
  );

  const stopped = await manager.stop();
  process.stderr.write(
    stopped
      ? "Stopped the shared Edge process.\n"
      : "The shared Edge process was not running.\n",
  );
}

main().catch((error: unknown) => {
  const logger = new StderrJsonLogger();
  logger.error("stop_shared_edge_failed", {
    errorType: error instanceof Error ? error.name : typeof error,
    message:
      error instanceof Error ? error.message : "Unknown stop failure",
  });
  process.exit(1);
});
