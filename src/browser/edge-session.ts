import { mkdir } from "node:fs/promises";
import { chromium, type Browser, type BrowserContext } from "playwright-core";

import type { AppConfig, EdgeMode } from "../config.js";
import type { Logger } from "../logger.js";
import { SharedEdgeManager } from "./shared-edge.js";

const CONNECT_ATTEMPTS = 3;
const CONNECT_RETRY_DELAY_MS = 500;

export class EdgeSession {
  private contextPromise: Promise<BrowserContext> | undefined;
  private sharedBrowser: Browser | undefined;
  private readonly edgeMode: EdgeMode;
  private readonly sharedEdge: SharedEdgeManager | undefined;

  constructor(
    private readonly config: AppConfig,
    private readonly logger: Logger,
    edgeModeOverride?: EdgeMode,
  ) {
    this.edgeMode = edgeModeOverride ?? config.edgeMode;
    if (this.edgeMode === "shared") {
      this.sharedEdge = new SharedEdgeManager(config, logger);
    }
  }

  async getContext(): Promise<BrowserContext> {
    if (!this.contextPromise) {
      this.contextPromise = this.launchContext();
    }

    try {
      return await this.contextPromise;
    } catch (error) {
      this.contextPromise = undefined;
      throw error;
    }
  }

  async close(): Promise<void> {
    const currentPromise = this.contextPromise;
    this.contextPromise = undefined;
    if (!currentPromise) {
      return;
    }

    try {
      const context = await currentPromise;
      if (this.edgeMode === "shared") {
        // Disconnect from the shared Edge only. The browser process is owned
        // by SharedEdgeManager and must keep running for other sessions.
        const browser = this.sharedBrowser;
        this.sharedBrowser = undefined;
        if (browser) {
          await browser.close();
        }
      } else {
        await context.close();
      }
    } catch (error) {
      this.logger.warn("edge_session_close_failed", {
        errorType: getErrorType(error),
      });
    }
  }

  private async launchContext(): Promise<BrowserContext> {
    if (this.edgeMode === "shared" && this.sharedEdge) {
      return this.connectToSharedContext(this.sharedEdge);
    }
    return this.launchPersistentContext();
  }

  private async connectToSharedContext(
    sharedEdge: SharedEdgeManager,
  ): Promise<BrowserContext> {
    const endpoint = await sharedEdge.ensureRunning();
    this.logger.info("edge_session_connecting_shared", {
      endpoint,
      headless: this.config.headless,
      siteOrigin: this.config.siteOrigin,
    });

    let lastError: unknown;
    for (let attempt = 1; attempt <= CONNECT_ATTEMPTS; attempt += 1) {
      try {
        const browser = await chromium.connectOverCDP(endpoint);
        const context = browser.contexts()[0];
        if (!context) {
          await browser.close().catch(() => undefined);
          throw new Error(
            "The shared Edge did not expose its default browser context.",
          );
        }

        this.sharedBrowser = browser;
        browser.once("disconnected", () => {
          if (this.sharedBrowser === browser) {
            this.sharedBrowser = undefined;
            this.contextPromise = undefined;
          }
          this.logger.info("edge_session_disconnected");
        });
        return context;
      } catch (error) {
        lastError = error;
        if (attempt < CONNECT_ATTEMPTS) {
          await new Promise((resolve) => {
            setTimeout(resolve, CONNECT_RETRY_DELAY_MS);
          });
        }
      }
    }

    throw lastError instanceof Error
      ? lastError
      : new Error("Failed to connect to the shared Edge process.");
  }

  private async launchPersistentContext(): Promise<BrowserContext> {
    await mkdir(this.config.profileDir, { recursive: true });
    this.logger.info("edge_session_starting", {
      browserChannel: this.config.browserChannel,
      headless: this.config.headless,
      siteOrigin: this.config.siteOrigin,
    });

    const context = await chromium.launchPersistentContext(this.config.profileDir, {
      channel: this.config.browserChannel,
      headless: this.config.headless,
      acceptDownloads: false,
      viewport: null,
    });

    context.once("close", () => {
      this.contextPromise = undefined;
      this.logger.info("edge_session_closed");
    });

    return context;
  }
}

function getErrorType(error: unknown): string {
  return error instanceof Error ? error.name : typeof error;
}
