import type {
  AuthStatus,
  AuthStatusService,
} from "./auth-status-service.js";
import {
  DEFAULT_FOLDER_RESULTS,
  MAX_DOWNLOAD_BYTES,
  type SharePointDocumentLibrariesResult,
  type SharePointDownloadedFile,
  type SharePointFolderResult,
} from "./file-content.js";
import type { SharePointFileService } from "./file-service.js";
import type { SharePointReadService } from "./read-service.js";
import {
  DEFAULT_SEARCH_RESULTS,
  type SharePointSearchOptions,
  type SharePointSearchResult,
} from "./search.js";

type AuthStatusProvider = Pick<AuthStatusService, "getStatus">;
type ReadProvider = Pick<
  SharePointReadService,
  "search" | "getPage" | "getListItem"
>;
type FileProvider = Pick<
  SharePointFileService,
  "listDocumentLibraries" | "listFolder" | "downloadFile"
>;

interface SiteService<T> {
  readonly siteUrl: string;
  readonly service: T;
}

export interface MultiSiteSearchOptions extends SharePointSearchOptions {
  readonly siteUrl?: string;
}

export interface MultiSiteSearchResult extends SharePointSearchResult {
  readonly siteUrls: readonly string[];
  readonly scopeUrls: readonly string[];
}

export interface MultiSiteDocumentLibrariesResult
  extends SharePointDocumentLibrariesResult {
  readonly siteUrls: readonly string[];
}

export interface MultiSiteAuthStatus extends AuthStatus {
  readonly siteUrls: readonly string[];
  readonly sites: readonly AuthStatus[];
}

export class MultiSiteAuthStatusService {
  constructor(
    private readonly entries: readonly SiteService<AuthStatusProvider>[],
  ) {
    assertSiteEntries(entries);
  }

  async getStatus(): Promise<MultiSiteAuthStatus> {
    const sites = await Promise.all(
      this.entries.map(({ service }) => service.getStatus()),
    );
    const primary = sites[0];
    if (!primary) {
      throw new Error("At least one SharePoint site must be configured.");
    }
    const failure = sites.find((status) => !status.authenticated);
    const authenticated = failure === undefined;
    const method = sites.some((status) => status.method === "browser-fetch")
      ? "browser-fetch"
      : primary.method;

    return {
      authenticated,
      state: failure?.state ?? "AUTHENTICATED",
      siteUrl: primary.siteUrl,
      siteUrls: sites.map((status) => status.siteUrl),
      sites,
      ...(method ? { method } : {}),
      ...(primary.user ? { user: primary.user } : {}),
      message: authenticated
        ? `The Edge session is authenticated for all ${sites.length} configured SharePoint sites.`
        : `SharePoint access failed for ${failure?.siteUrl ?? "a configured site"}: ${failure?.message ?? "Unknown error"}`,
    };
  }
}

export class MultiSiteSharePointReadService {
  constructor(
    private readonly entries: readonly SiteService<ReadProvider>[],
  ) {
    assertSiteEntries(entries);
  }

  async search(
    query: string,
    maxResults = DEFAULT_SEARCH_RESULTS,
    options: MultiSiteSearchOptions = {},
  ): Promise<MultiSiteSearchResult> {
    const { siteUrl, ...siteOptions } = options;
    const selected = siteUrl
      ? [selectEntry(this.entries, siteUrl)]
      : options.folderUrl
        ? [selectEntryForResource(this.entries, options.folderUrl)]
        : [...this.entries];

    if (selected.length > 1 && (options.startRow ?? 0) !== 0) {
      throw new Error(
        "Multi-site search paging requires siteUrl or folderUrl to select one configured site.",
      );
    }

    const results = await Promise.all(
      selected.map(({ service }) =>
        service.search(query, maxResults, siteOptions),
      ),
    );
    return combineSearchResults(results, maxResults);
  }

  async getPage(pageUrl: string) {
    return selectEntryForResource(this.entries, pageUrl).service.getPage(pageUrl);
  }

  async getListItem(itemUrl: string) {
    return selectEntryForResource(this.entries, itemUrl).service.getListItem(itemUrl);
  }
}

export class MultiSiteSharePointFileService {
  constructor(
    private readonly entries: readonly SiteService<FileProvider>[],
  ) {
    assertSiteEntries(entries);
  }

  async listDocumentLibraries(
    siteUrl?: string,
  ): Promise<MultiSiteDocumentLibrariesResult> {
    const selected = siteUrl
      ? [selectEntry(this.entries, siteUrl)]
      : [...this.entries];
    const results = await Promise.all(
      selected.map(({ service }) => service.listDocumentLibraries()),
    );
    const primary = results[0];
    if (!primary) {
      throw new Error("At least one SharePoint site must be configured.");
    }
    const libraries = results.flatMap((result) => result.libraries);
    return {
      siteUrl: primary.siteUrl,
      siteUrls: results.map((result) => result.siteUrl),
      returnedLibraries: libraries.length,
      method: results.some((result) => result.method === "browser-fetch")
        ? "browser-fetch"
        : "api-request",
      libraries,
    };
  }

  async listFolder(
    folderUrl: string,
    maxResults = DEFAULT_FOLDER_RESULTS,
  ): Promise<SharePointFolderResult> {
    return selectEntryForResource(this.entries, folderUrl).service.listFolder(
      folderUrl,
      maxResults,
    );
  }

  async downloadFile(
    fileUrl: string,
    maxBytes = MAX_DOWNLOAD_BYTES,
  ): Promise<SharePointDownloadedFile> {
    return selectEntryForResource(this.entries, fileUrl).service.downloadFile(
      fileUrl,
      maxBytes,
    );
  }
}

export function selectConfiguredSiteUrl(
  siteUrls: readonly string[],
  resourceUrl: string,
): string {
  const entries = siteUrls.map((siteUrl) => ({ siteUrl, service: undefined }));
  return selectEntryForResource(entries, resourceUrl).siteUrl;
}

function combineSearchResults(
  results: readonly SharePointSearchResult[],
  maxResults: number,
): MultiSiteSearchResult {
  const primary = results[0];
  if (!primary) {
    throw new Error("At least one SharePoint site must be configured.");
  }
  const unique = new Map<string, SharePointSearchResult["results"][number]>();
  for (const item of results.flatMap((result) => result.results)) {
    const key = item.url.toLowerCase();
    if (!unique.has(key)) {
      unique.set(key, item);
    }
  }
  const candidates = [...unique.values()].sort(
    (left, right) => (right.rank ?? 0) - (left.rank ?? 0),
  );
  const combined = candidates.slice(0, maxResults);
  const hasMore =
    candidates.length > combined.length || results.some((result) => result.hasMore);
  const { nextStartRow: _nextStartRow, ...primaryWithoutNextStartRow } = primary;

  return {
    ...primaryWithoutNextStartRow,
    siteUrls: results.map((result) => result.siteUrl),
    scopeUrls: results.map((result) => result.scopeUrl),
    totalRows: results.reduce((total, result) => total + result.totalRows, 0),
    returnedRows: combined.length,
    hasMore,
    ...(results.length === 1 && primary.nextStartRow !== undefined
      ? { nextStartRow: primary.nextStartRow }
      : {}),
    method: results.some((result) => result.method === "browser-fetch")
      ? "browser-fetch"
      : "api-request",
    results: combined,
  };
}

function selectEntry<T>(
  entries: readonly SiteService<T>[],
  requestedSiteUrl: string,
): SiteService<T> {
  let requested: URL;
  try {
    requested = new URL(requestedSiteUrl.trim());
  } catch {
    throw new Error("siteUrl must be an absolute configured SharePoint site URL.");
  }
  if (requested.search || requested.hash || requested.username || requested.password) {
    throw new Error("siteUrl must be an absolute configured SharePoint site URL.");
  }
  const key = canonicalSiteKey(requested);
  const selected = entries.find(
    (entry) => canonicalSiteKey(new URL(entry.siteUrl)) === key,
  );
  if (!selected) {
    throw new Error("siteUrl must identify one of the configured SharePoint sites.");
  }
  return selected;
}

function selectEntryForResource<T>(
  entries: readonly SiteService<T>[],
  resourceUrl: string,
): SiteService<T> {
  const primary = entries[0];
  if (!primary) {
    throw new Error("At least one SharePoint site must be configured.");
  }
  const raw = resourceUrl.trim();
  let candidate: URL;
  try {
    candidate = raw.startsWith("/")
      ? new URL(raw, new URL(primary.siteUrl).origin)
      : new URL(raw);
  } catch {
    throw new Error(
      "SharePoint resource URL must be absolute or server-relative.",
    );
  }
  if (
    candidate.protocol !== "https:" ||
    candidate.username ||
    candidate.password
  ) {
    throw new Error("SharePoint resource URL must use HTTPS without credentials.");
  }

  const selected = [...entries]
    .sort(
      (left, right) =>
        new URL(right.siteUrl).pathname.length -
        new URL(left.siteUrl).pathname.length,
    )
    .find((entry) => isResourceWithinSite(candidate, new URL(entry.siteUrl)));
  if (!selected) {
    throw new Error(
      "SharePoint resource URL must remain under one of the configured sites.",
    );
  }
  return selected;
}

function isResourceWithinSite(candidate: URL, site: URL): boolean {
  const sitePath = site.pathname.replace(/\/$/u, "").toLowerCase();
  const candidatePath = candidate.pathname.toLowerCase();
  return (
    candidate.origin === site.origin &&
    (candidatePath === sitePath || candidatePath.startsWith(`${sitePath}/`))
  );
}

function canonicalSiteKey(url: URL): string {
  return `${url.origin}${url.pathname.replace(/\/$/u, "")}`.toLowerCase();
}

function assertSiteEntries<T>(entries: readonly SiteService<T>[]): void {
  if (entries.length === 0) {
    throw new Error("At least one SharePoint site must be configured.");
  }
}
