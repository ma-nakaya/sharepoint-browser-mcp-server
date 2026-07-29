import { createHash } from "node:crypto";

import {
  buildDocumentLibrariesApiPath,
  buildFileDownloadApiPath,
  buildFileMetadataApiPath,
  buildFolderChildrenApiPaths,
  DEFAULT_FOLDER_RESULTS,
  MAX_DOWNLOAD_BYTES,
  normalizeFolderResultLimit,
  normalizeSharePointFileUrl,
  normalizeSharePointFolderUrl,
  parseDocumentLibrariesResponse,
  parseFileEntriesResponse,
  parseFileMetadataResponse,
  parseFolderEntriesResponse,
  type SharePointDocumentLibrariesResult,
  type SharePointDownloadedFile,
  type SharePointFolderResult,
} from "./file-content.js";
import type {
  RequestMethod,
  SharePointBinaryResponse,
  SharePointBinaryTransport,
  SharePointResponse,
} from "./http.js";

interface JsonResponse {
  readonly body: string;
  readonly method: RequestMethod;
}

export class SharePointFileService {
  constructor(
    private readonly siteUrl: string,
    private readonly transport: SharePointBinaryTransport,
  ) {}

  async listDocumentLibraries(): Promise<SharePointDocumentLibrariesResult> {
    const response = await this.getJson(
      buildDocumentLibrariesApiPath(),
      "document libraries",
    );
    const libraries = parseDocumentLibrariesResponse(response.body, this.siteUrl);
    return {
      siteUrl: this.siteUrl,
      returnedLibraries: libraries.length,
      method: response.method,
      libraries,
    };
  }

  async listFolder(
    folderUrl: string,
    maxResults = DEFAULT_FOLDER_RESULTS,
  ): Promise<SharePointFolderResult> {
    const folder = normalizeSharePointFolderUrl(this.siteUrl, folderUrl);
    const limit = normalizeFolderResultLimit(maxResults);
    const paths = buildFolderChildrenApiPaths(folder.serverRelativeUrl, limit);
    const [foldersResponse, filesResponse] = await Promise.all([
      this.getJson(paths.folders, "folder children"),
      this.getJson(paths.files, "folder files"),
    ]);
    const folders = parseFolderEntriesResponse(
      foldersResponse.body,
      this.siteUrl,
      folder,
    );
    const files = parseFileEntriesResponse(
      filesResponse.body,
      this.siteUrl,
      folder,
    );

    return {
      folderUrl: folder.url,
      serverRelativeUrl: folder.serverRelativeUrl,
      returnedFolders: folders.length,
      returnedFiles: files.length,
      limitReached: folders.length === limit || files.length === limit,
      method:
        foldersResponse.method === filesResponse.method
          ? foldersResponse.method
          : "mixed",
      folders,
      files,
    };
  }

  async downloadFile(fileUrl: string): Promise<SharePointDownloadedFile> {
    const file = normalizeSharePointFileUrl(this.siteUrl, fileUrl);
    const metadataResponse = await this.getJson(
      buildFileMetadataApiPath(file.serverRelativeUrl),
      "file metadata",
    );
    const metadata = parseFileMetadataResponse(
      metadataResponse.body,
      this.siteUrl,
      file,
    );
    if (metadata.sizeBytes > MAX_DOWNLOAD_BYTES) {
      throw new SharePointFileError(
        `File exceeds the ${MAX_DOWNLOAD_BYTES}-byte MCP download limit.`,
      );
    }

    const binaryResponse = await this.getBinary(
      buildFileDownloadApiPath(file.serverRelativeUrl),
      "file",
    );
    if (binaryResponse.body.byteLength !== metadata.sizeBytes) {
      throw new SharePointFileError(
        "Downloaded file size did not match SharePoint metadata.",
      );
    }
    const sha256 = createHash("sha256")
      .update(binaryResponse.body)
      .digest("hex");

    return {
      name: metadata.name,
      url: metadata.url,
      serverRelativeUrl: metadata.serverRelativeUrl,
      extension: metadata.extension,
      sizeBytes: metadata.sizeBytes,
      mimeType: metadata.mimeType,
      sha256,
      method: binaryResponse.method,
      data: binaryResponse.body,
    };
  }

  private async getJson(
    apiPath: string,
    resourceName: string,
  ): Promise<JsonResponse> {
    try {
      const primary = await this.transport.get(apiPath);
      const result = classifyJsonResponse(primary, resourceName, true);
      if (result) {
        return result;
      }
    } catch (error) {
      if (error instanceof SharePointFileError) {
        throw error;
      }
    }

    try {
      const fallback = await this.transport.getViaPage(apiPath);
      const result = classifyJsonResponse(fallback, resourceName, false);
      if (result) {
        return result;
      }
    } catch (error) {
      if (error instanceof SharePointFileError) {
        throw error;
      }
    }

    throw new SharePointFileError(
      `SharePoint ${resourceName} could not be read through the authenticated Edge session.`,
    );
  }

  private async getBinary(
    apiPath: string,
    resourceName: string,
  ): Promise<SharePointBinaryResponse> {
    try {
      const primary = await this.transport.getBinary(apiPath);
      const result = classifyBinaryResponse(primary, resourceName, true);
      if (result) {
        return result;
      }
    } catch (error) {
      if (error instanceof SharePointFileError) {
        throw error;
      }
    }

    try {
      const fallback = await this.transport.getBinaryViaPage(apiPath);
      const result = classifyBinaryResponse(fallback, resourceName, false);
      if (result) {
        return result;
      }
    } catch (error) {
      if (error instanceof SharePointFileError) {
        throw error;
      }
    }

    throw new SharePointFileError(
      `SharePoint ${resourceName} could not be downloaded through the authenticated Edge session.`,
    );
  }
}

class SharePointFileError extends Error {}

function classifyJsonResponse(
  response: SharePointResponse,
  resourceName: string,
  allowFallback: boolean,
): JsonResponse | undefined {
  if (response.status === 200) {
    if (response.bodyTruncated === true) {
      throw new SharePointFileError(
        `SharePoint ${resourceName} response exceeded the safe response limit.`,
      );
    }
    if (!isJsonContentType(response.contentType)) {
      if (allowFallback) {
        return undefined;
      }
      throw new SharePointFileError(
        `SharePoint returned a non-JSON ${resourceName} response.`,
      );
    }
    return { body: response.body, method: response.method };
  }
  return classifyErrorStatus(response.status, resourceName, allowFallback);
}

function classifyBinaryResponse(
  response: SharePointBinaryResponse,
  resourceName: string,
  allowFallback: boolean,
): SharePointBinaryResponse | undefined {
  if (response.status === 200) {
    if (response.bodyTruncated === true) {
      throw new SharePointFileError(
        `SharePoint ${resourceName} exceeded the safe download limit.`,
      );
    }
    if (response.contentType.toLowerCase().includes("text/html")) {
      if (allowFallback) {
        return undefined;
      }
      throw new SharePointFileError(
        "SharePoint login is required. Run npm run login with the dedicated Edge profile.",
      );
    }
    return response;
  }
  return classifyErrorStatus(response.status, resourceName, allowFallback);
}

function classifyErrorStatus<T>(
  status: number,
  resourceName: string,
  allowFallback: boolean,
): T | undefined {
  if (status === 0 || status === 401 || isRedirect(status)) {
    if (allowFallback) {
      return undefined;
    }
    throw new SharePointFileError(
      "SharePoint login is required. Run npm run login with the dedicated Edge profile.",
    );
  }
  if (status === 403) {
    if (allowFallback) {
      return undefined;
    }
    throw new SharePointFileError(
      `The signed-in user cannot read the requested SharePoint ${resourceName}.`,
    );
  }
  if (status === 404) {
    throw new SharePointFileError(
      `The requested SharePoint ${resourceName} was not found.`,
    );
  }
  throw new SharePointFileError(
    `SharePoint returned HTTP ${status} while reading ${resourceName}.`,
  );
}

function isJsonContentType(contentType: string): boolean {
  const normalized = contentType.toLowerCase();
  return normalized.includes("application/json") || normalized.includes("text/json");
}

function isRedirect(status: number): boolean {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}
