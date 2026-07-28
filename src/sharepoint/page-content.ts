import type { RequestMethod } from "./http.js";
import { encodeResourcePathArgument } from "./resource-path.js";

export const MAX_PAGE_TEXT_CHARACTERS = 50_000;

export interface NormalizedSharePointPage {
  readonly pageUrl: string;
  readonly serverRelativeUrl: string;
}

export interface SharePointPageContent {
  readonly title: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly modifiedTime?: string;
  readonly text: string;
  readonly truncated: boolean;
  readonly method: RequestMethod;
}

type JsonRecord = Record<string, unknown>;

export function normalizeSharePointPageUrl(
  siteUrl: string,
  pageUrl: string,
): NormalizedSharePointPage {
  const normalizedInput = pageUrl.trim();
  if (!normalizedInput) {
    throw new Error("Page URL must not be empty.");
  }

  const site = new URL(siteUrl);
  let candidate: URL;
  try {
    candidate = normalizedInput.startsWith("/")
      ? new URL(normalizedInput, site.origin)
      : new URL(normalizedInput);
  } catch {
    throw new Error("Page URL must be an absolute or server-relative URL.");
  }

  if (
    candidate.protocol !== "https:" ||
    candidate.origin !== site.origin ||
    candidate.username !== "" ||
    candidate.password !== ""
  ) {
    throw new Error("Page URL must remain on the configured SharePoint origin.");
  }

  let decodedPath: string;
  try {
    decodedPath = decodeURIComponent(candidate.pathname);
  } catch {
    throw new Error("Page URL contains invalid percent encoding.");
  }

  const sitePath = decodeURIComponent(site.pathname).replace(/\/$/u, "");
  const pagePrefix = `${sitePath}/SitePages/`;
  if (!decodedPath.toLowerCase().startsWith(pagePrefix.toLowerCase())) {
    throw new Error("Page URL must identify a page under the configured SitePages library.");
  }
  if (!decodedPath.toLowerCase().endsWith(".aspx")) {
    throw new Error("Page URL must identify an .aspx SharePoint page.");
  }

  const canonical = new URL(site.origin);
  canonical.pathname = decodedPath;

  return {
    pageUrl: canonical.toString(),
    serverRelativeUrl: decodedPath,
  };
}

export function buildPageContentApiPath(serverRelativeUrl: string): string {
  const encodedPath = encodeResourcePathArgument(serverRelativeUrl);
  const selectProperties = "Title,FileRef,CanvasContent1,Modified";
  return [
    `/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedPath}')`,
    `/ListItemAllFields?$select=${selectProperties}`,
  ].join("");
}

export function parsePageContentResponse(
  body: string,
  page: NormalizedSharePointPage,
  method: RequestMethod,
): SharePointPageContent {
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("SharePoint returned malformed page JSON.");
  }

  const root = asRecord(parsed);
  const item = asRecord(root?.d) ?? root;
  if (!item) {
    throw new Error("SharePoint returned an unexpected page response.");
  }

  const responsePath = asString(item.FileRef);
  if (
    !responsePath ||
    responsePath.toLowerCase() !== page.serverRelativeUrl.toLowerCase()
  ) {
    throw new Error("SharePoint returned page data for an unexpected path.");
  }

  const rawCanvas = asString(item.CanvasContent1) ?? "";
  const fullText = canvasHtmlToPlainText(rawCanvas);
  const truncated = fullText.length > MAX_PAGE_TEXT_CHARACTERS;
  const title =
    cleanText(asString(item.Title), 500) ??
    decodeLastPathSegment(page.serverRelativeUrl);
  const modifiedTime = cleanText(asString(item.Modified), 100);

  return {
    title,
    url: page.pageUrl,
    serverRelativeUrl: page.serverRelativeUrl,
    ...(modifiedTime ? { modifiedTime } : {}),
    text: truncated ? fullText.slice(0, MAX_PAGE_TEXT_CHARACTERS) : fullText,
    truncated,
    method,
  };
}

export function canvasHtmlToPlainText(value: string): string {
  const withoutNonContent = value
    .replace(/<!--[\s\S]*?-->/gu, " ")
    .replace(/<(script|style|noscript|template)\b[^>]*>[\s\S]*?<\/\1>/giu, " ");
  const withLineBreaks = withoutNonContent
    .replace(/<br\s*\/?>/giu, "\n")
    .replace(/<\/(p|div|li|h[1-6]|section|article|tr)>/giu, "\n");
  const withoutTags = withLineBreaks.replace(/<[^>]*>/gu, " ");

  return decodeHtmlEntities(withoutTags)
    .replace(/\r\n?/gu, "\n")
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/gu, " ").trim())
    .filter(Boolean)
    .join("\n");
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/giu, (_match, code: string) =>
      decodeNumericEntity(code, 16),
    )
    .replace(/&#([0-9]+);/gu, (_match, code: string) =>
      decodeNumericEntity(code, 10),
    )
    .replace(/&nbsp;/giu, " ")
    .replace(/&amp;/giu, "&")
    .replace(/&lt;/giu, "<")
    .replace(/&gt;/giu, ">")
    .replace(/&quot;/giu, '"')
    .replace(/&#39;|&apos;/giu, "'");
}

function decodeNumericEntity(value: string, radix: number): string {
  const codePoint = Number.parseInt(value, radix);
  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) {
    return "\uFFFD";
  }
  try {
    return String.fromCodePoint(codePoint);
  } catch {
    return "\uFFFD";
  }
}

function decodeLastPathSegment(pathname: string): string {
  return pathname.split("/").filter(Boolean).at(-1) ?? pathname;
}

function cleanText(value: string | undefined, maxCharacters: number): string | undefined {
  const normalized = value?.replace(/\s+/gu, " ").trim();
  if (!normalized) {
    return undefined;
  }
  return normalized.slice(0, maxCharacters);
}

function asRecord(value: unknown): JsonRecord | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}
