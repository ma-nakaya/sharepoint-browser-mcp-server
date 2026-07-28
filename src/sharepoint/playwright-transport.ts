import type { APIResponse, Page } from "playwright-core";

import type { AppConfig } from "../config.js";
import type { Logger } from "../logger.js";
import type { EdgeSession } from "../browser/edge-session.js";
import type { SharePointResponse, SharePointTransport } from "./http.js";
import { buildSharePointApiUrl } from "./url-guard.js";

const ACCEPT_JSON = "application/json;odata=nometadata";
const MAX_RESPONSE_CHARACTERS = 1_048_576;

export class PlaywrightSharePointTransport implements SharePointTransport {
  constructor(
    private readonly config: AppConfig,
    private readonly edgeSession: EdgeSession,
    private readonly logger: Logger,
  ) {}

  async get(apiPath: string): Promise<SharePointResponse> {
    const target = buildSharePointApiUrl(this.config.siteUrl, apiPath);
    const context = await this.edgeSession.getContext();
    const response = await context.request.get(target.toString(), {
      headers: { Accept: ACCEPT_JSON },
      timeout: this.config.requestTimeoutMs,
      failOnStatusCode: false,
      maxRedirects: 0,
    });

    return this.toSafeResponse(response);
  }

  async getViaPage(apiPath: string): Promise<SharePointResponse> {
    const target = buildSharePointApiUrl(this.config.siteUrl, apiPath);
    const context = await this.edgeSession.getContext();
    const page = await context.newPage();
    let keepPageOpenForLogin = false;

    try {
      await this.navigateToSite(page);
      const currentUrl = new URL(page.url());
      if (currentUrl.origin !== this.config.siteOrigin) {
        keepPageOpenForLogin = !this.config.headless;
        return {
          status: 401,
          contentType: "text/plain",
          body: "",
          method: "browser-fetch",
        };
      }

      const result = await page.evaluate(
        async ({ targetUrl, expectedOrigin, timeoutMs, maxResponseCharacters }) => {
          const controller = new AbortController();
          const timeout = setTimeout(() => controller.abort(), timeoutMs);
          try {
            const response = await fetch(targetUrl, {
              method: "GET",
              credentials: "include",
              headers: { Accept: "application/json;odata=nometadata" },
              redirect: "manual",
              signal: controller.signal,
            });
            if (response.type === "opaqueredirect") {
              return {
                status: 0,
                contentType: "",
                body: "",
                bodyTruncated: false,
                redirectedOutsideSite: true,
              };
            }

            const responseUrl = new URL(response.url || targetUrl);
            const contentType = response.headers.get("content-type") ?? "";
            const shouldReadBody =
              response.status === 200 && contentType.toLowerCase().includes("json");
            const responseBody = shouldReadBody ? await response.text() : "";
            return {
              status: response.status,
              contentType,
              body: responseBody.slice(0, maxResponseCharacters),
              bodyTruncated: responseBody.length > maxResponseCharacters,
              redirectedOutsideSite: responseUrl.origin !== expectedOrigin,
            };
          } finally {
            clearTimeout(timeout);
          }
        },
        {
          targetUrl: target.toString(),
          expectedOrigin: this.config.siteOrigin,
          timeoutMs: this.config.requestTimeoutMs,
          maxResponseCharacters: MAX_RESPONSE_CHARACTERS,
        },
      );

      return {
        status: result.redirectedOutsideSite ? 0 : result.status,
        contentType: result.contentType,
        body: result.body,
        bodyTruncated: result.bodyTruncated,
        method: "browser-fetch",
      };
    } finally {
      if (!keepPageOpenForLogin) {
        await page.close().catch((error: unknown) => {
          this.logger.warn("sharepoint_page_close_failed", {
            errorType: error instanceof Error ? error.name : typeof error,
          });
        });
      }
    }
  }

  async close(): Promise<void> {
    await this.edgeSession.close();
  }

  private async navigateToSite(page: Page): Promise<void> {
    try {
      await page.goto(this.config.siteUrl, {
        waitUntil: "domcontentloaded",
        timeout: this.config.requestTimeoutMs,
      });
    } catch (error) {
      this.logger.warn("sharepoint_site_navigation_incomplete", {
        errorType: error instanceof Error ? error.name : typeof error,
      });
      if (page.url() === "about:blank") {
        throw error;
      }
    }
  }

  private async toSafeResponse(response: APIResponse): Promise<SharePointResponse> {
    const status = response.status();
    const headers = response.headers();
    const contentType = headers["content-type"] ?? "";
    const shouldReadBody = status === 200 && contentType.toLowerCase().includes("json");
    const responseBody = shouldReadBody ? await response.text() : "";
    return {
      status,
      contentType,
      body: responseBody.slice(0, MAX_RESPONSE_CHARACTERS),
      bodyTruncated: responseBody.length > MAX_RESPONSE_CHARACTERS,
      method: "api-request",
    };
  }
}
