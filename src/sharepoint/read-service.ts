import type {
  RequestMethod,
  SharePointResponse,
  SharePointTransport,
} from "./http.js";
import {
  buildPageContentApiPath,
  normalizeSharePointPageUrl,
  parsePageContentResponse,
  type SharePointPageContent,
} from "./page-content.js";
import {
  buildSearchApiPath,
  DEFAULT_SEARCH_RESULTS,
  normalizeSearchQuery,
  normalizeSearchResultLimit,
  parseSearchResponse,
  type SharePointSearchResult,
} from "./search.js";

interface JsonResponse {
  readonly body: string;
  readonly method: RequestMethod;
}

export class SharePointReadService {
  constructor(
    private readonly siteUrl: string,
    private readonly transport: SharePointTransport,
  ) {}

  async search(
    query: string,
    maxResults = DEFAULT_SEARCH_RESULTS,
  ): Promise<SharePointSearchResult> {
    const normalizedQuery = normalizeSearchQuery(query);
    const normalizedLimit = normalizeSearchResultLimit(maxResults);
    const apiPath = buildSearchApiPath(
      this.siteUrl,
      normalizedQuery,
      normalizedLimit,
    );
    const response = await this.getJson(apiPath, "search");
    return parseSearchResponse(
      response.body,
      this.siteUrl,
      normalizedQuery,
      response.method,
    );
  }

  async getPage(pageUrl: string): Promise<SharePointPageContent> {
    const page = normalizeSharePointPageUrl(this.siteUrl, pageUrl);
    const apiPath = buildPageContentApiPath(page.serverRelativeUrl);
    const response = await this.getJson(apiPath, "page");
    return parsePageContentResponse(response.body, page, response.method);
  }

  private async getJson(apiPath: string, resourceName: string): Promise<JsonResponse> {
    try {
      const primary = await this.transport.get(apiPath);
      const primaryResult = classifyJsonResponse(primary, resourceName, true);
      if (primaryResult) {
        return primaryResult;
      }
    } catch (error) {
      if (error instanceof SharePointReadError) {
        throw error;
      }
      // Browser-page fetch is the fallback for session/proxy differences.
    }

    try {
      const fallback = await this.transport.getViaPage(apiPath);
      const fallbackResult = classifyJsonResponse(fallback, resourceName, false);
      if (fallbackResult) {
        return fallbackResult;
      }
    } catch (error) {
      if (error instanceof SharePointReadError) {
        throw error;
      }
    }

    throw new SharePointReadError(
      `SharePoint ${resourceName} data could not be read through the authenticated Edge session.`,
    );
  }
}

class SharePointReadError extends Error {}

function classifyJsonResponse(
  response: SharePointResponse,
  resourceName: string,
  allowFallback: boolean,
): JsonResponse | undefined {
  if (response.status === 200) {
    if (response.bodyTruncated === true) {
      throw new SharePointReadError(
        `SharePoint ${resourceName} response exceeded the safe response limit.`,
      );
    }
    if (!isJsonContentType(response.contentType)) {
      if (allowFallback) {
        return undefined;
      }
      throw new SharePointReadError(
        `SharePoint returned a non-JSON ${resourceName} response.`,
      );
    }
    return { body: response.body, method: response.method };
  }

  if (
    response.status === 0 ||
    response.status === 401 ||
    isRedirect(response.status)
  ) {
    if (allowFallback) {
      return undefined;
    }
    throw new SharePointReadError(
      "SharePoint login is required. Run npm run login with the dedicated Edge profile.",
    );
  }

  if (response.status === 403) {
    throw new SharePointReadError(
      `The signed-in user cannot read the requested SharePoint ${resourceName}.`,
    );
  }
  if (response.status === 404) {
    throw new SharePointReadError(
      `The requested SharePoint ${resourceName} was not found.`,
    );
  }

  throw new SharePointReadError(
    `SharePoint returned HTTP ${response.status} while reading ${resourceName}.`,
  );
}

function isJsonContentType(contentType: string): boolean {
  const normalized = contentType.toLowerCase();
  return normalized.includes("application/json") || normalized.includes("text/json");
}

function isRedirect(status: number): boolean {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}
