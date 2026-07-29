import {
  extractDocumentNodes,
  extractDocumentOutline,
  extractDocumentText,
  searchDocumentStructure,
  type DocumentOutlineNode,
  type DocumentSearchMatch,
  type ExtractedDocumentNode,
  type DocumentUnitType,
  type SupportedDocumentFormat,
} from "./document-extractor.js";
import { MAX_DOCUMENT_SOURCE_BYTES } from "./file-content.js";
import type { SharePointFileService } from "./file-service.js";
import type { RequestMethod } from "./http.js";

export interface SharePointDocumentText {
  readonly name: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly sizeBytes: number;
  readonly mimeType: string;
  readonly sha256: string;
  readonly method: RequestMethod;
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly text: string;
  readonly characters: number;
  readonly truncated: boolean;
}

interface SharePointDocumentMetadata {
  readonly name: string;
  readonly url: string;
  readonly serverRelativeUrl: string;
  readonly sizeBytes: number;
  readonly mimeType: string;
  readonly sha256: string;
  readonly method: RequestMethod;
}

export interface SharePointDocumentOutline extends SharePointDocumentMetadata {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly nodeCount: number;
  readonly truncated: boolean;
  readonly nodes: readonly DocumentOutlineNode[];
}

export interface SharePointDocumentNodes extends SharePointDocumentMetadata {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly requestedNodeIds: readonly string[];
  readonly returnedNodes: number;
  readonly characters: number;
  readonly truncated: boolean;
  readonly nodes: readonly ExtractedDocumentNode[];
}

export interface SharePointDocumentSearch extends SharePointDocumentMetadata {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly query: string;
  readonly matchedNodes: number;
  readonly returnedNodes: number;
  readonly truncated: boolean;
  readonly results: readonly DocumentSearchMatch[];
}

export class SharePointDocumentService {
  constructor(private readonly fileService: SharePointFileService) {}

  async extractText(fileUrl: string): Promise<SharePointDocumentText> {
    const file = await this.fileService.downloadFile(
      fileUrl,
      MAX_DOCUMENT_SOURCE_BYTES,
    );
    const extracted = await extractDocumentText(file.data, file.extension);
    return {
      ...getDocumentMetadata(file),
      ...extracted,
    };
  }

  async getOutline(fileUrl: string): Promise<SharePointDocumentOutline> {
    const file = await this.fileService.downloadFile(
      fileUrl,
      MAX_DOCUMENT_SOURCE_BYTES,
    );
    const outline = await extractDocumentOutline(file.data, file.extension);
    return {
      ...getDocumentMetadata(file),
      ...outline,
    };
  }

  async getNodes(
    fileUrl: string,
    nodeIds: readonly string[],
    expectedSha256?: string,
  ): Promise<SharePointDocumentNodes> {
    const file = await this.fileService.downloadFile(
      fileUrl,
      MAX_DOCUMENT_SOURCE_BYTES,
    );
    assertExpectedDocumentHash(file.sha256, expectedSha256);
    const selected = await extractDocumentNodes(
      file.data,
      file.extension,
      nodeIds,
    );
    return {
      ...getDocumentMetadata(file),
      ...selected,
    };
  }

  async search(
    fileUrl: string,
    query: string,
    maxResults?: number,
    expectedSha256?: string,
  ): Promise<SharePointDocumentSearch> {
    const file = await this.fileService.downloadFile(
      fileUrl,
      MAX_DOCUMENT_SOURCE_BYTES,
    );
    assertExpectedDocumentHash(file.sha256, expectedSha256);
    const result = await searchDocumentStructure(
      file.data,
      file.extension,
      query,
      maxResults,
    );
    return {
      ...getDocumentMetadata(file),
      ...result,
    };
  }
}

function getDocumentMetadata(
  file: Awaited<ReturnType<SharePointFileService["downloadFile"]>>,
): SharePointDocumentMetadata {
  return {
    name: file.name,
    url: file.url,
    serverRelativeUrl: file.serverRelativeUrl,
    sizeBytes: file.sizeBytes,
    mimeType: file.mimeType,
    sha256: file.sha256,
    method: file.method,
  };
}

function assertExpectedDocumentHash(
  actualSha256: string,
  expectedSha256: string | undefined,
): void {
  if (
    expectedSha256 !== undefined &&
    actualSha256 !== expectedSha256.toLowerCase()
  ) {
    throw new Error(
      "The SharePoint document changed after its outline or search result was created. Get a new outline before selecting nodes.",
    );
  }
}
