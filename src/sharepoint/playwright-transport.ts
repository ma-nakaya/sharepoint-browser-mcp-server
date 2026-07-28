import type { APIResponse, Page } from "playwright-core";

import type { AppConfig } from "../config.js";
import type { Logger } from "../logger.js";
import type { EdgeSession } from "../browser/edge-session.js";
import type {
  SharePointBinaryResponse,
  SharePointBinaryTransport,
  SharePointResponse,
} from "./http.js";
import { buildSharePointApiUrl } from "./url-guard.js";

const ACCEPT_JSON = "application/json;odata=nometadata";
const MAX_RESPONSE_CHARACTERS = 1_048_576;
const MAX_BINARY_RESPONSE_BYTES = 5 * 1_024 * 1_024;

export class PlaywrightSharePointTransport implements SharePointBinaryTransport {
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

  async getBinary(apiPath: string): Promise<SharePointBinaryResponse> {
    const target = buildSharePointApiUrl(this.config.siteUrl, apiPath);
    const context = await this.edgeSession.getContext();
    const response = await context.request.get(target.toString(), {
      headers: { Accept: "*/*" },
      timeout: this.config.requestTimeoutMs,
      failOnStatusCode: false,
      maxRedirects: 0,
    });

    const status = response.status();
    const headers = response.headers();
    const contentType = headers["content-type"] ?? "";
    const declaredLength = parseContentLength(headers["content-length"]);
    if (status !== 200) {
      return {
        status,
        contentType,
        body: new Uint8Array(),
        bodyTruncated: false,
        method: "api-request",
      };
    }
    if (declaredLength !== undefined && declaredLength > MAX_BINARY_RESPONSE_BYTES) {
      return {
        status,
        contentType,
        body: new Uint8Array(),
        bodyTruncated: true,
        method: "api-request",
      };
    }

    const responseBody = await response.body();
    const bodyTruncated = responseBody.byteLength > MAX_BINARY_RESPONSE_BYTES;
    return {
      status,
      contentType,
      body: bodyTruncated ? new Uint8Array() : responseBody,
      bodyTruncated,
      method: "api-request",
    };
  }

  async getBinaryViaPage(apiPath: string): Promise<SharePointBinaryResponse> {
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
          body: new Uint8Array(),
          bodyTruncated: false,
          method: "browser-fetch",
        };
      }

      const result = await page.evaluate(
        async ({ targetUrl, expectedOrigin, timeoutMs, maxResponseBytes }) => {
          const controller = new AbortController();
          const timeout = setTimeout(() => controller.abort(), timeoutMs);
          try {
            const response = await fetch(targetUrl, {
              method: "GET",
              credentials: "include",
              headers: { Accept: "*/*" },
              redirect: "manual",
              signal: controller.signal,
            });
            if (response.type === "opaqueredirect") {
              return {
                status: 0,
                contentType: "",
                bodyBase64: "",
                bodyTruncated: false,
                redirectedOutsideSite: true,
              };
            }

            const responseUrl = new URL(response.url || targetUrl);
            const contentType = response.headers.get("content-type") ?? "";
            const rawLength = response.headers.get("content-length");
            const declaredLength = rawLength === null ? undefined : Number(rawLength);
            if (
              response.status === 200 &&
              declaredLength !== undefined &&
              Number.isSafeInteger(declaredLength) &&
              declaredLength > maxResponseBytes
            ) {
              return {
                status: response.status,
                contentType,
                bodyBase64: "",
                bodyTruncated: true,
                redirectedOutsideSite: responseUrl.origin !== expectedOrigin,
              };
            }

            const bytes =
              response.status === 200
                ? new Uint8Array(await response.arrayBuffer())
                : new Uint8Array();
            if (bytes.byteLength > maxResponseBytes) {
              return {
                status: response.status,
                contentType,
                bodyBase64: "",
                bodyTruncated: true,
                redirectedOutsideSite: responseUrl.origin !== expectedOrigin,
              };
            }

            let binary = "";
            for (let offset = 0; offset < bytes.length; offset += 32_768) {
              binary += String.fromCharCode(
                ...bytes.subarray(offset, offset + 32_768),
              );
            }
            return {
              status: response.status,
              contentType,
              bodyBase64: btoa(binary),
              bodyTruncated: false,
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
          maxResponseBytes: MAX_BINARY_RESPONSE_BYTES,
        },
      );

      return {
        status: result.redirectedOutsideSite ? 0 : result.status,
        contentType: result.contentType,
        body: result.bodyTruncated
          ? new Uint8Array()
          : Buffer.from(result.bodyBase64, "base64"),
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

function parseContentLength(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : undefined;
}
