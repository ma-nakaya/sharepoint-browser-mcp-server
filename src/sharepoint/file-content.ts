import type { RequestMethod } from "./http.js";
import { encodeResourcePathArgument } from "./resource-path.js";

export const DEFAULT_FOLDER_RESULTS = 50;
export const MAX_FOLDER_RESULTS = 100;
export const MAX_DOWNLOAD_BYTES = 5 * 1_024 * 1_024;
export const MAX_DOCUMENT_SOURCE_BYTES = 20 * 1_024 * 1_024;

const DOWNLOAD_MIME_TYPES = new Map<string, string>([
  [".bmp", "image/bmp"],
  [".csv", "text/csv"],
  [".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
  [".gif", "image/gif"],
  [".jpeg", "image/jpeg"],
  [".jpg", "image/jpeg"],
  [".json", "application/json"],
  [".md", "text/markdown"],
  [".pdf", "application/pdf"],
  [".png", "image/png"],
  [".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"],
  [".txt", "text/plain"],
  [".webp", "image/webp"],
  [".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
  [".xml", "application/xml"],
]);

export interface NormalizedSharePointResource {
  readonly url: string;
  readonly serverRelativeUrl: string;
}

export interface SharePointDocumentLibrary {
  readonly title: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly itemCount: number;
}

export interface SharePointDocumentLibrariesResult {
  readonly siteUrl: string;
  readonly returnedLibraries: number;
  readonly method: RequestMethod;
  readonly libraries: readonly SharePointDocumentLibrary[];
}

export interface SharePointFolderEntry {
  readonly name: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly itemCount: number;
}

export interface SharePointFileEntry {
  readonly name: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly extension: string;
  readonly sizeBytes: number;
  readonly modifiedTime?: string;
  readonly majorVersion?: number;
  readonly downloadable: boolean;
}

export interface SharePointFolderResult {
  readonly folderUrl: string;
  readonly serverRelativeUrl: string;
  readonly returnedFolders: number;
  readonly returnedFiles: number;
  readonly limitReached: boolean;
  readonly method: RequestMethod | "mixed";
  readonly folders: readonly SharePointFolderEntry[];
  readonly files: readonly SharePointFileEntry[];
}

export interface SharePointFileMetadata extends SharePointFileEntry {
  readonly mimeType: string;
}

export interface SharePointDownloadedFile {
  readonly name: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly extension: string;
  readonly sizeBytes: number;
  readonly mimeType: string;
  readonly sha256: string;
  readonly method: RequestMethod;
  readonly data: Uint8Array;
}

type JsonRecord = Record<string, unknown>;

export function normalizeSharePointFolderUrl(
  siteUrl: string,
  folderUrl: string,
): NormalizedSharePointResource {
  return normalizeSiteResourceUrl(siteUrl, folderUrl, "folder");
}

export function normalizeSharePointFileUrl(
  siteUrl: string,
  fileUrl: string,
): NormalizedSharePointResource {
  const file = normalizeSiteResourceUrl(siteUrl, fileUrl, "file");
  const extension = getFileExtension(file.serverRelativeUrl);
  if (!DOWNLOAD_MIME_TYPES.has(extension)) {
    throw new Error(
      `File type ${extension || "(none)"} is not allowed for MCP download.`,
    );
  }
  return file;
}

export function normalizeFolderResultLimit(value: number): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > MAX_FOLDER_RESULTS) {
    throw new Error(`Folder result limit must be between 1 and ${MAX_FOLDER_RESULTS}.`);
  }
  return value;
}

export function buildDocumentLibrariesApiPath(): string {
  const select = "Title,ItemCount,RootFolder/ServerRelativeUrl";
  return [
    `/_api/web/lists?$select=${select}`,
    "&$expand=RootFolder",
    "&$filter=BaseTemplate%20eq%20101%20and%20Hidden%20eq%20false",
    "&$orderby=Title",
    "&$top=100",
  ].join("");
}

export function buildFolderChildrenApiPaths(
  serverRelativeUrl: string,
  maxResults: number,
): { readonly folders: string; readonly files: string } {
  const limit = normalizeFolderResultLimit(maxResults);
  const encodedPath = encodeResourcePathArgument(serverRelativeUrl);
  const base = `/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodedPath}')`;
  return {
    folders: [
      `${base}/Folders?$select=Name,ServerRelativeUrl,ItemCount`,
      "&$orderby=Name",
      `&$top=${limit}`,
    ].join(""),
    files: [
      `${base}/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified,MajorVersion`,
      "&$orderby=Name",
      `&$top=${limit}`,
    ].join(""),
  };
}

export function buildFileMetadataApiPath(serverRelativeUrl: string): string {
  const encodedPath = encodeResourcePathArgument(serverRelativeUrl);
  return [
    `/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedPath}')`,
    "?$select=Name,ServerRelativeUrl,Length,TimeLastModified,MajorVersion",
  ].join("");
}

export function buildFileDownloadApiPath(serverRelativeUrl: string): string {
  const encodedPath = encodeResourcePathArgument(serverRelativeUrl);
  return `/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedPath}')/$value`;
}

export function parseDocumentLibrariesResponse(
  body: string,
  siteUrl: string,
): SharePointDocumentLibrary[] {
  const values = parseCollectionBody(body, "document libraries");
  return values
    .map((value) => {
      const record = asRecord(value);
      const rootFolder = asRecord(record?.RootFolder);
      const title = cleanText(asString(record?.Title), 500);
      const serverRelativeUrl = asString(rootFolder?.ServerRelativeUrl);
      const itemCount = asNonNegativeInteger(record?.ItemCount);
      if (!title || !serverRelativeUrl || itemCount === undefined) {
        return undefined;
      }
      const normalized = tryNormalizeSiteResourceUrl(
        siteUrl,
        serverRelativeUrl,
        "folder",
      );
      return normalized
        ? {
            title,
            url: normalized.url,
            serverRelativeUrl: normalized.serverRelativeUrl,
            itemCount,
          }
        : undefined;
    })
    .filter((value): value is SharePointDocumentLibrary => value !== undefined);
}

export function parseFolderEntriesResponse(
  body: string,
  siteUrl: string,
  folder: NormalizedSharePointResource,
): SharePointFolderEntry[] {
  const values = parseCollectionBody(body, "folders");
  return values
    .map((value) => {
      const record = asRecord(value);
      const name = cleanText(asString(record?.Name), 500);
      const serverRelativeUrl = asString(record?.ServerRelativeUrl);
      const itemCount = asNonNegativeInteger(record?.ItemCount);
      if (!name || !serverRelativeUrl || itemCount === undefined) {
        return undefined;
      }
      const normalized = tryNormalizeSiteResourceUrl(
        siteUrl,
        serverRelativeUrl,
        "folder",
      );
      return normalized && isDirectChild(normalized.serverRelativeUrl, folder.serverRelativeUrl)
        ? {
            name,
            url: normalized.url,
            serverRelativeUrl: normalized.serverRelativeUrl,
            itemCount,
          }
        : undefined;
    })
    .filter((value): value is SharePointFolderEntry => value !== undefined);
}

export function parseFileEntriesResponse(
  body: string,
  siteUrl: string,
  folder: NormalizedSharePointResource,
): SharePointFileEntry[] {
  const values = parseCollectionBody(body, "files");
  return values
    .map((value) => parseFileRecord(value, siteUrl, folder.serverRelativeUrl))
    .filter((value): value is SharePointFileEntry => value !== undefined);
}

export function parseFileMetadataResponse(
  body: string,
  siteUrl: string,
  expectedFile: NormalizedSharePointResource,
): SharePointFileMetadata {
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("SharePoint returned malformed file metadata JSON.");
  }
  const root = asRecord(parsed);
  const record = asRecord(root?.d) ?? root;
  const file = parseFileRecord(record, siteUrl);
  if (
    !file ||
    file.serverRelativeUrl.toLowerCase() !==
      expectedFile.serverRelativeUrl.toLowerCase()
  ) {
    throw new Error("SharePoint returned metadata for an unexpected file.");
  }
  if (!DOWNLOAD_MIME_TYPES.has(file.extension)) {
    throw new Error("SharePoint returned a file type that is not allowed for download.");
  }
  return {
    ...file,
    mimeType: DOWNLOAD_MIME_TYPES.get(file.extension) ?? "application/octet-stream",
  };
}

function parseFileRecord(
  value: unknown,
  siteUrl: string,
  expectedParent?: string,
): SharePointFileEntry | undefined {
  const record = asRecord(value);
  const name = cleanText(asString(record?.Name), 500);
  const serverRelativeUrl = asString(record?.ServerRelativeUrl);
  const sizeBytes = asNonNegativeInteger(record?.Length);
  if (!name || !serverRelativeUrl || sizeBytes === undefined) {
    return undefined;
  }
  const normalized = tryNormalizeSiteResourceUrl(siteUrl, serverRelativeUrl, "file");
  if (
    !normalized ||
    (expectedParent &&
      !isDirectChild(normalized.serverRelativeUrl, expectedParent))
  ) {
    return undefined;
  }
  const extension = getFileExtension(normalized.serverRelativeUrl);
  const modifiedTime = cleanText(asString(record?.TimeLastModified), 100);
  const majorVersion = asNonNegativeInteger(record?.MajorVersion);
  return {
    name,
    url: normalized.url,
    serverRelativeUrl: normalized.serverRelativeUrl,
    extension,
    sizeBytes,
    ...(modifiedTime ? { modifiedTime } : {}),
    ...(majorVersion !== undefined ? { majorVersion } : {}),
    downloadable:
      DOWNLOAD_MIME_TYPES.has(extension) && sizeBytes <= MAX_DOWNLOAD_BYTES,
  };
}

function normalizeSiteResourceUrl(
  siteUrl: string,
  resourceUrl: string,
  kind: "file" | "folder",
): NormalizedSharePointResource {
  const input = resourceUrl.trim();
  if (!input) {
    throw new Error(`${capitalize(kind)} URL must not be empty.`);
  }
  const site = new URL(siteUrl);
  let candidate: URL;
  try {
    candidate = input.startsWith("/")
      ? new URL(input, site.origin)
      : new URL(input);
  } catch {
    throw new Error(
      `${capitalize(kind)} URL must be an absolute or server-relative URL.`,
    );
  }
  if (
    candidate.protocol !== "https:" ||
    candidate.origin !== site.origin ||
    candidate.username !== "" ||
    candidate.password !== ""
  ) {
    throw new Error(
      `${capitalize(kind)} URL must remain on the configured SharePoint origin.`,
    );
  }

  let decodedPath: string;
  try {
    decodedPath = decodeURIComponent(candidate.pathname);
  } catch {
    throw new Error(`${capitalize(kind)} URL contains invalid percent encoding.`);
  }
  const sitePath = decodeURIComponent(site.pathname).replace(/\/$/u, "");
  const withinSite = decodedPath
    .toLowerCase()
    .startsWith(`${sitePath.toLowerCase()}/`);
  if (
    !withinSite ||
    decodedPath.includes("\0") ||
    decodedPath.includes("\\") ||
    decodedPath.endsWith("/")
  ) {
    throw new Error(
      `${capitalize(kind)} URL must identify an item under the configured SharePoint site.`,
    );
  }

  const canonical = new URL(site.origin);
  canonical.pathname = decodedPath;
  return {
    url: canonical.toString(),
    serverRelativeUrl: decodedPath,
  };
}

function tryNormalizeSiteResourceUrl(
  siteUrl: string,
  resourceUrl: string,
  kind: "file" | "folder",
): NormalizedSharePointResource | undefined {
  try {
    return normalizeSiteResourceUrl(siteUrl, resourceUrl, kind);
  } catch {
    return undefined;
  }
}

function parseCollectionBody(body: string, resourceName: string): unknown[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error(`SharePoint returned malformed ${resourceName} JSON.`);
  }
  if (Array.isArray(parsed)) {
    return parsed;
  }
  const root = asRecord(parsed);
  if (Array.isArray(root?.value)) {
    return root.value;
  }
  const verbose = asRecord(root?.d);
  if (Array.isArray(verbose?.results)) {
    return verbose.results;
  }
  throw new Error(`SharePoint returned an unexpected ${resourceName} response.`);
}

function isDirectChild(candidate: string, parent: string): boolean {
  const normalizedParent = parent.replace(/\/$/u, "").toLowerCase();
  const normalizedCandidate = candidate.toLowerCase();
  if (!normalizedCandidate.startsWith(`${normalizedParent}/`)) {
    return false;
  }
  return !normalizedCandidate.slice(normalizedParent.length + 1).includes("/");
}

function getFileExtension(value: string): string {
  const name = value.split("/").at(-1) ?? "";
  const dotIndex = name.lastIndexOf(".");
  return dotIndex > 0 ? name.slice(dotIndex).toLowerCase() : "";
}

function cleanText(value: string | undefined, maxCharacters: number): string | undefined {
  const normalized = value?.replace(/\s+/gu, " ").trim();
  return normalized ? normalized.slice(0, maxCharacters) : undefined;
}

function asRecord(value: unknown): JsonRecord | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function asNonNegativeInteger(value: unknown): number | undefined {
  const parsed = typeof value === "number" ? value : Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function capitalize(value: string): string {
  return `${value.slice(0, 1).toUpperCase()}${value.slice(1)}`;
}
