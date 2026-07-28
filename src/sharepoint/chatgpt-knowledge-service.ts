import type { SharePointDocumentService } from "./document-service.js";
import type { SharePointReadService } from "./read-service.js";
import { MAX_SEARCH_RESULTS } from "./search.js";

export const CHATGPT_SEARCH_RESULTS = 10;

const SUPPORTED_DOCUMENT_EXTENSIONS = new Set([
  "pdf",
  "docx",
  "xlsx",
  "pptx",
]);

export interface ChatGptSearchItem {
  readonly id: string;
  readonly title: string;
  readonly url: string;
}

export interface ChatGptSearchResult {
  readonly results: readonly ChatGptSearchItem[];
}

export type ChatGptFetchMetadataValue = string | number | boolean;

export interface ChatGptFetchResult {
  readonly id: string;
  readonly title: string;
  readonly text: string;
  readonly url: string;
  readonly metadata: Readonly<Record<string, ChatGptFetchMetadataValue>>;
}

type ReadService = Pick<SharePointReadService, "search" | "getPage">;
type DocumentService = Pick<SharePointDocumentService, "extractText">;

export class ChatGptKnowledgeService {
  constructor(
    private readonly readService: ReadService,
    private readonly documentService: DocumentService,
  ) {}

  async search(query: string): Promise<ChatGptSearchResult> {
    const result = await this.readService.search(query, MAX_SEARCH_RESULTS);
    const results = result.results
      .filter((item) => isFetchableSearchItem(item.kind, item.fileExtension))
      .slice(0, CHATGPT_SEARCH_RESULTS)
      .map((item) => ({
        id: item.url,
        title: item.title,
        url: item.url,
      }));

    return { results };
  }

  async fetch(id: string): Promise<ChatGptFetchResult> {
    const normalizedId = id.trim();
    if (!normalizedId) {
      throw new Error("Knowledge item ID must not be empty.");
    }

    const extension = getPathExtension(normalizedId);
    if (extension === "aspx") {
      const page = await this.readService.getPage(normalizedId);
      return {
        id: page.url,
        title: page.title,
        text: page.text,
        url: page.url,
        metadata: {
          source: "sharepoint",
          kind: "page",
          serverRelativeUrl: page.serverRelativeUrl,
          ...(page.modifiedTime ? { modifiedTime: page.modifiedTime } : {}),
          truncated: page.truncated,
        },
      };
    }

    if (extension && SUPPORTED_DOCUMENT_EXTENSIONS.has(extension)) {
      const document = await this.documentService.extractText(normalizedId);
      return {
        id: document.url,
        title: document.name,
        text: document.text,
        url: document.url,
        metadata: {
          source: "sharepoint",
          kind: "document",
          format: document.format,
          unitType: document.unitType,
          unitCount: document.unitCount,
          sizeBytes: document.sizeBytes,
          mimeType: document.mimeType,
          sha256: document.sha256,
          truncated: document.truncated,
        },
      };
    }

    throw new Error(
      "Knowledge item ID must identify a SharePoint page or a supported PDF, DOCX, XLSX, or PPTX document.",
    );
  }
}

function isFetchableSearchItem(
  kind: "page" | "document" | "file" | "other",
  fileExtension: string | undefined,
): boolean {
  if (kind === "page") {
    return fileExtension?.toLowerCase() === "aspx";
  }
  return (
    kind === "document" &&
    fileExtension !== undefined &&
    SUPPORTED_DOCUMENT_EXTENSIONS.has(fileExtension.toLowerCase())
  );
}

function getPathExtension(value: string): string | undefined {
  try {
    const pathname = new URL(value).pathname;
    const fileName = pathname.split("/").at(-1) ?? "";
    const extension = fileName.includes(".") ? fileName.split(".").at(-1) : undefined;
    return extension?.toLowerCase();
  } catch {
    if (!value.startsWith("/")) {
      return undefined;
    }
    const pathname = value.split(/[?#]/u, 1)[0] ?? "";
    const fileName = pathname.split("/").at(-1) ?? "";
    const extension = fileName.includes(".") ? fileName.split(".").at(-1) : undefined;
    return extension?.toLowerCase();
  }
}
