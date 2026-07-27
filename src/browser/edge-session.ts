import { mkdir } from "node:fs/promises";
import { chromium, type BrowserContext } from "playwright-core";

import type { AppConfig } from "../config.js";
import type { Logger } from "../logger.js";

export class EdgeSession {
  private contextPromise: Promise<BrowserContext> | undefined;

  constructor(
    private readonly config: AppConfig,
    private readonly logger: Logger,
  ) {}

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
      await context.close();
    } catch (error) {
      this.logger.warn("edge_session_close_failed", {
        errorType: getErrorType(error),
      });
    }
  }

  private async launchContext(): Promise<BrowserContext> {
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
