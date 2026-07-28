import {
  extractDocumentText,
  type DocumentUnitType,
  type SupportedDocumentFormat,
} from "./document-extractor.js";
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

export class SharePointDocumentService {
  constructor(private readonly fileService: SharePointFileService) {}

  async extractText(fileUrl: string): Promise<SharePointDocumentText> {
    const file = await this.fileService.downloadFile(fileUrl);
    const extracted = await extractDocumentText(file.data, file.extension);
    return {
      name: file.name,
      url: file.url,
      serverRelativeUrl: file.serverRelativeUrl,
      sizeBytes: file.sizeBytes,
      mimeType: file.mimeType,
      sha256: file.sha256,
      method: file.method,
      ...extracted,
    };
  }
}
