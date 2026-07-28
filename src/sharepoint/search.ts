import type { RequestMethod } from "./http.js";
import { normalizeSharePointFolderUrl } from "./file-content.js";

export const DEFAULT_SEARCH_RESULTS = 10;
export const MAX_SEARCH_RESULTS = 20;
export const MAX_SEARCH_QUERY_CHARACTERS = 200;
export const MAX_SEARCH_START_ROW = 50_000;
export const MAX_SEARCH_FILE_EXTENSIONS = 10;

export type SharePointSearchScope = "all" | "pages" | "documents";
export type SharePointSearchSort = "relevance" | "modified-desc";

export interface SharePointSearchOptions {
  readonly startRow?: number;
  readonly scope?: SharePointSearchScope;
  readonly folderUrl?: string;
  readonly fileExtensions?: readonly string[];
  readonly modifiedAfter?: string;
  readonly modifiedBefore?: string;
  readonly sort?: SharePointSearchSort;
}

interface NormalizedSharePointSearchOptions {
  readonly maxResults: number;
  readonly startRow: number;
  readonly scope: SharePointSearchScope;
  readonly scopeUrl: string;
  readonly fileExtensions: readonly string[];
  readonly modifiedAfter?: string;
  readonly modifiedBefore?: string;
  readonly sort: SharePointSearchSort;
}

const SEARCH_PROPERTIES = [
  "Title",
  "Path",
  "ParentLink",
  "Author",
  "Size",
  "Rank",
  "FileExtension",
  "ContentClass",
  "LastModifiedTime",
  "HitHighlightedSummary",
] as const;

export type SharePointSearchItemKind = "page" | "document" | "file" | "other";

export interface SharePointSearchItem {
  readonly title: string;
  readonly url: string;
  readonly kind: SharePointSearchItemKind;
  readonly parentUrl?: string;
  readonly author?: string;
  readonly sizeBytes?: number;
  readonly rank?: number;
  readonly fileExtension?: string;
  readonly contentClass?: string;
  readonly modifiedTime?: string;
  readonly summary?: string;
}

export interface SharePointSearchResult {
  readonly query: string;
  readonly siteUrl: string;
  readonly totalRows: number;
  readonly returnedRows: number;
  readonly startRow: number;
  readonly hasMore: boolean;
  readonly nextStartRow?: number;
  readonly scope: SharePointSearchScope;
  readonly scopeUrl: string;
  readonly fileExtensions: readonly string[];
  readonly modifiedAfter?: string;
  readonly modifiedBefore?: string;
  readonly sort: SharePointSearchSort;
  readonly method: RequestMethod;
  readonly results: readonly SharePointSearchItem[];
}

type JsonRecord = Record<string, unknown>;

export function normalizeSearchQuery(query: string): string {
  const normalized = query.trim();
  if (normalized.length === 0) {
    throw new Error("Search query must not be empty.");
  }
  if (normalized.length > MAX_SEARCH_QUERY_CHARACTERS) {
    throw new Error(
      `Search query must be ${MAX_SEARCH_QUERY_CHARACTERS} characters or fewer.`,
    );
  }
  return normalized;
}

export function normalizeSearchResultLimit(value: number): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > MAX_SEARCH_RESULTS) {
    throw new Error(`Search result limit must be between 1 and ${MAX_SEARCH_RESULTS}.`);
  }
  return value;
}

export function normalizeSearchStartRow(value: number): number {
  if (!Number.isSafeInteger(value) || value < 0 || value > MAX_SEARCH_START_ROW) {
    throw new Error(
      `Search start row must be between 0 and ${MAX_SEARCH_START_ROW}.`,
    );
  }
  return value;
}

export function buildSearchApiPath(
  siteUrl: string,
  query: string,
  maxResults: number,
  options: SharePointSearchOptions = {},
): string {
  const normalizedQuery = normalizeSearchQuery(query);
  const normalized = normalizeSearchOptions(siteUrl, maxResults, options);
  const queryTemplate = buildQueryTemplate(normalized);
  const parameters = new URLSearchParams({
    querytext: `'${escapeODataString(normalizedQuery)}'`,
    querytemplate: `'${escapeODataString(queryTemplate)}'`,
    rowlimit: String(normalized.maxResults),
    rowsperpage: String(normalized.maxResults),
    startrow: String(normalized.startRow),
    trimduplicates: "true",
    enablestemming: "true",
    selectproperties: `'${SEARCH_PROPERTIES.join(",")}'`,
  });
  if (normalized.sort === "modified-desc") {
    parameters.set("sortlist", "'LastModifiedTime:descending'");
  }

  return `/_api/search/query?${parameters.toString()}`;
}

export function parseSearchResponse(
  body: string,
  siteUrl: string,
  query: string,
  method: RequestMethod,
  maxResults = DEFAULT_SEARCH_RESULTS,
  options: SharePointSearchOptions = {},
): SharePointSearchResult {
  const normalized = normalizeSearchOptions(siteUrl, maxResults, options);
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("SharePoint returned malformed search JSON.");
  }

  const root = asRecord(parsed);
  if (!root) {
    throw new Error("SharePoint returned an unexpected search response.");
  }
  const verboseRoot = asRecord(root.d);
  const queryResult =
    asRecord(verboseRoot?.query) ??
    asRecord(root.query) ??
    (root.PrimaryQueryResult ? root : undefined);
  const relevantResults = asRecord(
    asRecord(queryResult?.PrimaryQueryResult)?.RelevantResults,
  );
  if (!relevantResults) {
    throw new Error("SharePoint returned an unexpected search response.");
  }

  const rows = unwrapResults(asRecord(relevantResults.Table)?.Rows);
  const results = rows
    .map((row) => parseSearchRow(row, siteUrl))
    .filter((item): item is SharePointSearchItem => item !== undefined);
  const totalRows = asNonNegativeInteger(relevantResults.TotalRows) ?? results.length;
  const nextStartRow = normalized.startRow + normalized.maxResults;
  const hasMore = nextStartRow < totalRows;

  return {
    query: normalizeSearchQuery(query),
    siteUrl,
    totalRows,
    returnedRows: results.length,
    startRow: normalized.startRow,
    hasMore,
    ...(hasMore ? { nextStartRow } : {}),
    scope: normalized.scope,
    scopeUrl: normalized.scopeUrl,
    fileExtensions: normalized.fileExtensions,
    ...(normalized.modifiedAfter
      ? { modifiedAfter: normalized.modifiedAfter }
      : {}),
    ...(normalized.modifiedBefore
      ? { modifiedBefore: normalized.modifiedBefore }
      : {}),
    sort: normalized.sort,
    method,
    results,
  };
}

function parseSearchRow(value: unknown, siteUrl: string): SharePointSearchItem | undefined {
  const row = asRecord(value);
  const cells = unwrapResults(row?.Cells);
  const properties = new Map<string, string>();

  for (const cellValue of cells) {
    const cell = asRecord(cellValue);
    const key = asString(cell?.Key);
    const cellText = asString(cell?.Value);
    if (key && cellText !== undefined) {
      properties.set(key, cellText);
    }
  }

  const rawPath = properties.get("Path");
  if (!rawPath) {
    return undefined;
  }
  const scopedUrl = normalizeScopedResultUrl(rawPath, siteUrl);
  if (!scopedUrl) {
    return undefined;
  }

  const title =
    cleanText(properties.get("Title"), 500) ??
    decodeLastPathSegment(new URL(scopedUrl).pathname);
  const fileExtension = cleanText(properties.get("FileExtension"), 32)?.toLowerCase();
  const parentUrl = normalizeOptionalScopedResultUrl(
    properties.get("ParentLink"),
    siteUrl,
  );
  const author = cleanText(properties.get("Author"), 500);
  const sizeBytes = asNonNegativeInteger(properties.get("Size"));
  const rank = asFiniteNumber(properties.get("Rank"));
  const contentClass = cleanText(properties.get("ContentClass"), 200);
  const modifiedTime = cleanText(properties.get("LastModifiedTime"), 100);
  const summary = cleanText(
    htmlToPlainText(properties.get("HitHighlightedSummary") ?? ""),
    2_000,
  );

  return {
    title,
    url: scopedUrl,
    kind: classifySearchItem(fileExtension, contentClass),
    ...(parentUrl ? { parentUrl } : {}),
    ...(author ? { author } : {}),
    ...(sizeBytes !== undefined ? { sizeBytes } : {}),
    ...(rank !== undefined ? { rank } : {}),
    ...(fileExtension ? { fileExtension } : {}),
    ...(contentClass ? { contentClass } : {}),
    ...(modifiedTime ? { modifiedTime } : {}),
    ...(summary ? { summary } : {}),
  };
}

function normalizeSearchOptions(
  siteUrl: string,
  maxResults: number,
  options: SharePointSearchOptions,
): NormalizedSharePointSearchOptions {
  const scope = options.scope ?? "all";
  const startRow = normalizeSearchStartRow(options.startRow ?? 0);
  const scopeUrl = options.folderUrl
    ? normalizeSharePointFolderUrl(siteUrl, options.folderUrl).url.replace(/\/$/u, "")
    : siteUrl.replace(/\/$/u, "");
  const fileExtensions = normalizeFileExtensions(options.fileExtensions ?? []);
  const modifiedAfter = normalizeDateOnly(options.modifiedAfter, "modifiedAfter");
  const modifiedBefore = normalizeDateOnly(options.modifiedBefore, "modifiedBefore");
  const sort = options.sort ?? "relevance";

  if (scope === "pages" && fileExtensions.length > 0) {
    throw new Error("File extension filters cannot be combined with page-only search.");
  }
  if (
    modifiedAfter &&
    modifiedBefore &&
    modifiedAfter.localeCompare(modifiedBefore, "en") > 0
  ) {
    throw new Error("modifiedAfter must not be later than modifiedBefore.");
  }

  return {
    maxResults: normalizeSearchResultLimit(maxResults),
    startRow,
    scope,
    scopeUrl,
    fileExtensions,
    ...(modifiedAfter ? { modifiedAfter } : {}),
    ...(modifiedBefore ? { modifiedBefore } : {}),
    sort,
  };
}

function buildQueryTemplate(options: NormalizedSharePointSearchOptions): string {
  const constraints = [
    "({searchterms})",
    `Path:"${options.scopeUrl}/*"`,
  ];

  if (options.scope === "pages") {
    constraints.push("FileType:aspx");
  } else if (options.scope === "documents") {
    constraints.push("IsDocument:True", "NOT FileType:aspx");
  }

  if (options.fileExtensions.length === 1) {
    constraints.push(`FileType:${options.fileExtensions[0]}`);
  } else if (options.fileExtensions.length > 1) {
    constraints.push(
      `(${options.fileExtensions
        .map((extension) => `FileType:${extension}`)
        .join(" OR ")})`,
    );
  }

  if (options.modifiedAfter) {
    constraints.push(`LastModifiedTime>=${options.modifiedAfter}`);
  }
  if (options.modifiedBefore) {
    constraints.push(`LastModifiedTime<=${options.modifiedBefore}`);
  }

  return constraints.join(" AND ");
}

function normalizeFileExtensions(values: readonly string[]): string[] {
  if (values.length > MAX_SEARCH_FILE_EXTENSIONS) {
    throw new Error(
      `At most ${MAX_SEARCH_FILE_EXTENSIONS} file extensions may be filtered.`,
    );
  }
  const normalized = values.map((value) =>
    value.trim().replace(/^\./u, "").toLowerCase(),
  );
  if (normalized.some((value) => !/^[a-z0-9]{1,10}$/u.test(value))) {
    throw new Error("File extensions must contain only letters and numbers.");
  }
  return [...new Set(normalized)];
}

function normalizeDateOnly(
  value: string | undefined,
  fieldName: string,
): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  const normalized = value.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/u.test(normalized)) {
    throw new Error(`${fieldName} must use YYYY-MM-DD format.`);
  }
  const parsed = new Date(`${normalized}T00:00:00.000Z`);
  if (
    Number.isNaN(parsed.getTime()) ||
    parsed.toISOString().slice(0, 10) !== normalized
  ) {
    throw new Error(`${fieldName} must be a valid calendar date.`);
  }
  return normalized;
}

function classifySearchItem(
  fileExtension: string | undefined,
  contentClass: string | undefined,
): SharePointSearchItemKind {
  if (fileExtension === "aspx") {
    return "page";
  }
  if (
    fileExtension &&
    ["pdf", "docx", "xlsx", "pptx"].includes(fileExtension)
  ) {
    return "document";
  }
  if (fileExtension) {
    return "file";
  }
  return contentClass?.toLowerCase().includes("document") ? "document" : "other";
}

function normalizeOptionalScopedResultUrl(
  value: string | undefined,
  siteUrl: string,
): string | undefined {
  return value ? normalizeScopedResultUrl(value, siteUrl) : undefined;
}

function normalizeScopedResultUrl(value: string, siteUrl: string): string | undefined {
  try {
    const site = new URL(siteUrl);
    const candidate = new URL(value, site.origin);
    const sitePath = site.pathname.replace(/\/$/u, "").toLowerCase();
    const candidatePath = candidate.pathname.toLowerCase();
    const withinSite =
      candidatePath === sitePath || candidatePath.startsWith(`${sitePath}/`);

    if (
      candidate.protocol !== "https:" ||
      candidate.origin !== site.origin ||
      !withinSite
    ) {
      return undefined;
    }

    candidate.hash = "";
    return candidate.toString();
  } catch {
    return undefined;
  }
}

function htmlToPlainText(value: string): string {
  return decodeHtmlEntities(value.replace(/<[^>]*>/gu, " "))
    .replace(/\s+/gu, " ")
    .trim();
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&nbsp;/giu, " ")
    .replace(/&amp;/giu, "&")
    .replace(/&lt;/giu, "<")
    .replace(/&gt;/giu, ">")
    .replace(/&quot;/giu, '"')
    .replace(/&#39;|&apos;/giu, "'");
}

function decodeLastPathSegment(pathname: string): string {
  const segment = pathname.split("/").filter(Boolean).at(-1) ?? pathname;
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function unwrapResults(value: unknown): unknown[] {
  if (Array.isArray(value)) {
    return value;
  }
  const record = asRecord(value);
  return Array.isArray(record?.results) ? record.results : [];
}

function asRecord(value: unknown): JsonRecord | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function cleanText(value: string | undefined, maxCharacters: number): string | undefined {
  const normalized = value?.replace(/\s+/gu, " ").trim();
  if (!normalized) {
    return undefined;
  }
  return normalized.slice(0, maxCharacters);
}

function asNonNegativeInteger(value: unknown): number | undefined {
  const parsed = typeof value === "number" ? value : Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function asFiniteNumber(value: unknown): number | undefined {
  const parsed = typeof value === "number" ? value : Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function escapeODataString(value: string): string {
  return value.replaceAll("'", "''");
}
