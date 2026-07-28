import type { RequestMethod } from "./http.js";

export const DEFAULT_SEARCH_RESULTS = 10;
export const MAX_SEARCH_RESULTS = 20;
export const MAX_SEARCH_QUERY_CHARACTERS = 200;

const SEARCH_PROPERTIES = [
  "Title",
  "Path",
  "FileExtension",
  "ContentClass",
  "LastModifiedTime",
  "HitHighlightedSummary",
] as const;

export interface SharePointSearchItem {
  readonly title: string;
  readonly url: string;
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

export function buildSearchApiPath(
  siteUrl: string,
  query: string,
  maxResults: number,
): string {
  const normalizedQuery = normalizeSearchQuery(query);
  const normalizedLimit = normalizeSearchResultLimit(maxResults);
  const queryTemplate = `({searchterms}) AND Path:"${siteUrl}/*"`;
  const parameters = new URLSearchParams({
    querytext: `'${escapeODataString(normalizedQuery)}'`,
    querytemplate: `'${escapeODataString(queryTemplate)}'`,
    rowlimit: String(normalizedLimit),
    trimduplicates: "true",
    selectproperties: `'${SEARCH_PROPERTIES.join(",")}'`,
  });

  return `/_api/search/query?${parameters.toString()}`;
}

export function parseSearchResponse(
  body: string,
  siteUrl: string,
  query: string,
  method: RequestMethod,
): SharePointSearchResult {
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

  return {
    query: normalizeSearchQuery(query),
    siteUrl,
    totalRows,
    returnedRows: results.length,
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
  const fileExtension = cleanText(properties.get("FileExtension"), 32);
  const contentClass = cleanText(properties.get("ContentClass"), 200);
  const modifiedTime = cleanText(properties.get("LastModifiedTime"), 100);
  const summary = cleanText(
    htmlToPlainText(properties.get("HitHighlightedSummary") ?? ""),
    2_000,
  );

  return {
    title,
    url: scopedUrl,
    ...(fileExtension ? { fileExtension } : {}),
    ...(contentClass ? { contentClass } : {}),
    ...(modifiedTime ? { modifiedTime } : {}),
    ...(summary ? { summary } : {}),
  };
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

function escapeODataString(value: string): string {
  return value.replaceAll("'", "''");
}
