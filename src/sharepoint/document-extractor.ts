import path from "node:path";

import { strFromU8, unzipSync, type UnzipFileInfo } from "fflate";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

export const MAX_EXTRACTED_TEXT_CHARACTERS = 100_000;
export const MAX_PDF_PAGES = 200;

const MAX_ARCHIVE_ENTRIES = 1_000;
const MAX_ARCHIVE_PART_BYTES = 2 * 1_024 * 1_024;
const MAX_ARCHIVE_SELECTED_BYTES = 8 * 1_024 * 1_024;
const MAX_OFFICE_UNITS = 200;
const MAX_SPREADSHEET_CELLS = 20_000;

export type SupportedDocumentFormat = "pdf" | "docx" | "xlsx" | "pptx";
export type DocumentUnitType = "pages" | "parts" | "sheets" | "slides";

export interface ExtractedDocumentText {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly text: string;
  readonly characters: number;
  readonly truncated: boolean;
}

interface SpreadsheetSheet {
  readonly name: string;
  readonly path: string;
}

export async function extractDocumentText(
  data: Uint8Array,
  extension: string,
): Promise<ExtractedDocumentText> {
  switch (extension.toLowerCase()) {
    case ".pdf":
      return extractPdfText(data);
    case ".docx":
      return extractDocxText(data);
    case ".xlsx":
      return extractXlsxText(data);
    case ".pptx":
      return extractPptxText(data);
    default:
      throw new Error("Document type is not supported for text extraction.");
  }
}

async function extractPdfText(data: Uint8Array): Promise<ExtractedDocumentText> {
  const loadingTask = getDocument({
    data: new Uint8Array(data),
    disableFontFace: true,
    enableXfa: false,
    useSystemFonts: false,
    useWasm: false,
    verbosity: 0,
  });
  let document;
  try {
    document = await loadingTask.promise;
    if (document.numPages > MAX_PDF_PAGES) {
      throw new Error(`PDF exceeds the ${MAX_PDF_PAGES}-page extraction limit.`);
    }

    const output = new TextAccumulator();
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await document.getPage(pageNumber);
      const content = await page.getTextContent();
      let pageText = "";
      for (const item of content.items) {
        if (!("str" in item)) {
          continue;
        }
        pageText += item.str;
        pageText += item.hasEOL ? "\n" : " ";
      }
      output.append(`[Page ${pageNumber}]\n${pageText}`);
      page.cleanup();
      if (output.truncated) {
        break;
      }
    }

    return output.result("pdf", "pages", document.numPages);
  } catch (error) {
    if (error instanceof Error && /exceeds the \d+-page/u.test(error.message)) {
      throw error;
    }
    throw new Error("PDF text could not be extracted.");
  } finally {
    if (document) {
      await document.cleanup().catch(() => undefined);
    }
    await loadingTask.destroy().catch(() => undefined);
  }
}

function extractDocxText(data: Uint8Array): ExtractedDocumentText {
  const files = unzipSelectedXml(data, (name) =>
    name === "word/document.xml" ||
    /^word\/(?:header|footer)\d+\.xml$/u.test(name) ||
    /^word\/(?:footnotes|endnotes)\.xml$/u.test(name),
  );
  if (!files["word/document.xml"]) {
    throw new Error("DOCX archive does not contain word/document.xml.");
  }

  const orderedParts = Object.keys(files).sort((left, right) => {
    if (left === "word/document.xml") {
      return -1;
    }
    if (right === "word/document.xml") {
      return 1;
    }
    return naturalCompare(left, right);
  });
  if (orderedParts.length > MAX_OFFICE_UNITS) {
    throw new Error(`DOCX exceeds the ${MAX_OFFICE_UNITS}-part extraction limit.`);
  }

  const output = new TextAccumulator();
  for (const partName of orderedParts) {
    const xml = readXmlPart(files, partName);
    const label =
      partName === "word/document.xml"
        ? "Document"
        : partName.replace(/^word\//u, "").replace(/\.xml$/u, "");
    output.append(`[${label}]\n${extractWordprocessingText(xml)}`);
    if (output.truncated) {
      break;
    }
  }
  return output.result("docx", "parts", orderedParts.length);
}

function extractXlsxText(data: Uint8Array): ExtractedDocumentText {
  const files = unzipSelectedXml(data, (name) =>
    name === "xl/workbook.xml" ||
    name === "xl/_rels/workbook.xml.rels" ||
    name === "xl/sharedStrings.xml" ||
    /^xl\/worksheets\/sheet\d+\.xml$/u.test(name),
  );
  const worksheetPaths = Object.keys(files)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/u.test(name))
    .sort(naturalCompare);
  if (worksheetPaths.length === 0) {
    throw new Error("XLSX archive does not contain any worksheets.");
  }
  if (worksheetPaths.length > MAX_OFFICE_UNITS) {
    throw new Error(`XLSX exceeds the ${MAX_OFFICE_UNITS}-sheet extraction limit.`);
  }

  const sharedStrings = files["xl/sharedStrings.xml"]
    ? parseSharedStrings(readXmlPart(files, "xl/sharedStrings.xml"))
    : [];
  const sheets = resolveSpreadsheetSheets(files, worksheetPaths);
  const output = new TextAccumulator();
  let cellsRemaining = MAX_SPREADSHEET_CELLS;

  for (const sheet of sheets) {
    const sheetResult = extractWorksheetText(
      readXmlPart(files, sheet.path),
      sharedStrings,
      cellsRemaining,
    );
    cellsRemaining -= sheetResult.cells;
    output.append(`[Sheet: ${sheet.name}]\n${sheetResult.text}`);
    if (sheetResult.limitReached) {
      output.markTruncated();
    }
    if (output.truncated) {
      break;
    }
  }

  return output.result("xlsx", "sheets", sheets.length);
}

function extractPptxText(data: Uint8Array): ExtractedDocumentText {
  const files = unzipSelectedXml(data, (name) =>
    /^ppt\/slides\/slide\d+\.xml$/u.test(name) ||
    /^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(name),
  );
  const slidePaths = Object.keys(files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name))
    .sort(naturalCompare);
  if (slidePaths.length === 0) {
    throw new Error("PPTX archive does not contain any slides.");
  }
  if (slidePaths.length > MAX_OFFICE_UNITS) {
    throw new Error(`PPTX exceeds the ${MAX_OFFICE_UNITS}-slide extraction limit.`);
  }

  const output = new TextAccumulator();
  for (let index = 0; index < slidePaths.length; index += 1) {
    const slideNumber = index + 1;
    const slidePath = slidePaths[index];
    if (!slidePath) {
      continue;
    }
    output.append(
      `[Slide ${slideNumber}]\n${extractPresentationText(
        readXmlPart(files, slidePath),
      )}`,
    );
    const notesPath = `ppt/notesSlides/notesSlide${slideNumber}.xml`;
    if (files[notesPath]) {
      const notes = extractPresentationText(readXmlPart(files, notesPath));
      if (notes) {
        output.append(`[Notes ${slideNumber}]\n${notes}`);
      }
    }
    if (output.truncated) {
      break;
    }
  }

  return output.result("pptx", "slides", slidePaths.length);
}

function unzipSelectedXml(
  data: Uint8Array,
  select: (name: string) => boolean,
): Record<string, Uint8Array> {
  let entries = 0;
  let selectedBytes = 0;
  try {
    const files = unzipSync(data, {
      filter(file: UnzipFileInfo): boolean {
        entries += 1;
        if (entries > MAX_ARCHIVE_ENTRIES) {
          throw new Error("Office archive contains too many entries.");
        }
        if (!isSafeArchivePath(file.name) || !select(file.name)) {
          return false;
        }
        if (file.originalSize > MAX_ARCHIVE_PART_BYTES) {
          throw new Error("Office XML part exceeds the safe extraction limit.");
        }
        selectedBytes += file.originalSize;
        if (selectedBytes > MAX_ARCHIVE_SELECTED_BYTES) {
          throw new Error("Office archive exceeds the safe expanded-size limit.");
        }
        return true;
      },
    });
    let actualBytes = 0;
    for (const bytes of Object.values(files)) {
      if (bytes.byteLength > MAX_ARCHIVE_PART_BYTES) {
        throw new Error("Office XML part exceeds the safe extraction limit.");
      }
      actualBytes += bytes.byteLength;
      if (actualBytes > MAX_ARCHIVE_SELECTED_BYTES) {
        throw new Error("Office archive exceeds the safe expanded-size limit.");
      }
    }
    return files;
  } catch (error) {
    if (error instanceof Error && /^Office /u.test(error.message)) {
      throw error;
    }
    throw new Error("Office Open XML archive could not be read.");
  }
}

function readXmlPart(
  files: Record<string, Uint8Array>,
  name: string,
): string {
  const bytes = files[name];
  if (!bytes) {
    throw new Error(`Office archive is missing required part: ${name}.`);
  }
  const xml = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/iu.test(xml)) {
    throw new Error("Office XML document type declarations are not allowed.");
  }
  return xml;
}

function extractWordprocessingText(xml: string): string {
  return xmlToPlainText(
    xml
      .replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/giu, " ")
      .replace(/<w:instrText\b[^>]*>[\s\S]*?<\/w:instrText>/giu, " ")
      .replace(/<w:(?:br|cr)\b[^>]*\/?>/giu, "\n")
      .replace(/<w:tab\b[^>]*\/?>/giu, "\t")
      .replace(/<\/w:tc>/giu, "\t")
      .replace(/<\/w:p>/giu, "\n"),
  );
}

function extractPresentationText(xml: string): string {
  return xmlToPlainText(
    xml
      .replace(/<a:br\b[^>]*\/?>/giu, "\n")
      .replace(/<a:tab\b[^>]*\/?>/giu, "\t")
      .replace(/<\/a:p>/giu, "\n"),
  );
}

function parseSharedStrings(xml: string): string[] {
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/giu)].map((match) =>
    extractTextElements(match[1] ?? ""),
  );
}

function resolveSpreadsheetSheets(
  files: Record<string, Uint8Array>,
  fallbackPaths: string[],
): SpreadsheetSheet[] {
  const workbookBytes = files["xl/workbook.xml"];
  const relationshipsBytes = files["xl/_rels/workbook.xml.rels"];
  if (!workbookBytes || !relationshipsBytes) {
    return fallbackPaths.map((sheetPath, index) => ({
      name: `Sheet ${index + 1}`,
      path: sheetPath,
    }));
  }

  const workbookXml = readXmlPart(files, "xl/workbook.xml");
  const relationshipsXml = readXmlPart(
    files,
    "xl/_rels/workbook.xml.rels",
  );
  const relationships = new Map<string, string>();
  for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/giu)) {
    const attributes = parseXmlAttributes(match[1] ?? "");
    const id = attributes.get("Id");
    const target = attributes.get("Target");
    if (!id || !target || target.includes("\\") || target.includes("\0")) {
      continue;
    }
    const normalized = path.posix.normalize(
      target.startsWith("/") ? target.slice(1) : `xl/${target}`,
    );
    if (normalized.startsWith("../") || !files[normalized]) {
      continue;
    }
    relationships.set(id, normalized);
  }

  const sheets: SpreadsheetSheet[] = [];
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/giu)) {
    const attributes = parseXmlAttributes(match[1] ?? "");
    const name = attributes.get("name");
    const relationshipId = attributes.get("r:id");
    const sheetPath = relationshipId
      ? relationships.get(relationshipId)
      : undefined;
    if (name && sheetPath && /^xl\/worksheets\/sheet\d+\.xml$/u.test(sheetPath)) {
      sheets.push({ name: cleanInlineText(name), path: sheetPath });
    }
  }
  return sheets.length > 0
    ? sheets
    : fallbackPaths.map((sheetPath, index) => ({
        name: `Sheet ${index + 1}`,
        path: sheetPath,
      }));
}

function extractWorksheetText(
  xml: string,
  sharedStrings: string[],
  cellLimit: number,
): { readonly text: string; readonly cells: number; readonly limitReached: boolean } {
  const lines: string[] = [];
  let cells = 0;
  let limitReached = false;
  for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/giu)) {
    if (cells >= cellLimit) {
      limitReached = true;
      break;
    }
    const attributes = parseXmlAttributes(match[1] ?? "");
    const reference = attributes.get("r") ?? `Cell ${cells + 1}`;
    const type = attributes.get("t") ?? "";
    const body = match[2] ?? "";
    const rawValue = firstXmlElementText(body, "v");
    let value: string | undefined;
    if (type === "s" && rawValue !== undefined) {
      const index = Number(rawValue);
      value = Number.isSafeInteger(index) && index >= 0
        ? sharedStrings[index]
        : undefined;
    } else if (type === "inlineStr") {
      value = extractTextElements(body);
    } else if (type === "b" && rawValue !== undefined) {
      value = rawValue === "1" ? "TRUE" : "FALSE";
    } else if (type === "e" && rawValue !== undefined) {
      value = `#ERROR ${decodeXmlEntities(rawValue)}`;
    } else if (rawValue !== undefined) {
      value = decodeXmlEntities(rawValue);
    }
    const normalizedValue = cleanInlineText(value ?? "");
    if (!normalizedValue) {
      continue;
    }
    lines.push(`${reference}: ${normalizedValue}`);
    cells += 1;
  }
  return { text: lines.join("\n"), cells, limitReached };
}

function firstXmlElementText(xml: string, localName: string): string | undefined {
  const pattern = new RegExp(
    `<(?:[\\w.-]+:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${localName}>`,
    "iu",
  );
  return pattern.exec(xml)?.[1];
}

function extractTextElements(xml: string): string {
  return [...xml.matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/giu)]
    .map((match) => decodeXmlEntities(stripXmlTags(match[1] ?? "")))
    .join("");
}

function parseXmlAttributes(value: string): Map<string, string> {
  const attributes = new Map<string, string>();
  const pattern = /([:\w.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/gu;
  for (const match of value.matchAll(pattern)) {
    const name = match[1];
    const rawValue = match[2] ?? match[3];
    if (name && rawValue !== undefined) {
      attributes.set(name, decodeXmlEntities(rawValue));
    }
  }
  return attributes;
}

function xmlToPlainText(xml: string): string {
  return normalizeOutputText(decodeXmlEntities(stripXmlTags(xml)));
}

function stripXmlTags(value: string): string {
  return value
    .replace(/<!--[\s\S]*?-->/gu, " ")
    .replace(/<\?[\s\S]*?\?>/gu, " ")
    .replace(/<[^>]*>/gu, " ");
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/giu, (_match, code: string) =>
      decodeCodePoint(code, 16),
    )
    .replace(/&#([0-9]+);/gu, (_match, code: string) =>
      decodeCodePoint(code, 10),
    )
    .replace(/&lt;/giu, "<")
    .replace(/&gt;/giu, ">")
    .replace(/&quot;/giu, '"')
    .replace(/&apos;/giu, "'")
    .replace(/&amp;/giu, "&");
}

function decodeCodePoint(value: string, radix: number): string {
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

function normalizeOutputText(value: string): string {
  return value
    .replace(/\r\n?/gu, "\n")
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/gu, " ").trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

function cleanInlineText(value: string): string {
  return value.replace(/\s+/gu, " ").trim();
}

function isSafeArchivePath(value: string): boolean {
  return (
    value.length > 0 &&
    !value.startsWith("/") &&
    !value.includes("\\") &&
    !value.includes("\0") &&
    !value.split("/").includes("..")
  );
}

function naturalCompare(left: string, right: string): number {
  return left.localeCompare(right, "en", {
    numeric: true,
    sensitivity: "base",
  });
}

class TextAccumulator {
  private value = "";
  public truncated = false;

  append(value: string): void {
    const normalized = normalizeOutputText(value);
    if (!normalized || this.truncated) {
      return;
    }
    const separator = this.value ? "\n\n" : "";
    const available =
      MAX_EXTRACTED_TEXT_CHARACTERS - this.value.length - separator.length;
    if (available <= 0) {
      this.truncated = true;
      return;
    }
    if (normalized.length > available) {
      this.value += `${separator}${normalized.slice(0, available)}`;
      this.truncated = true;
      return;
    }
    this.value += `${separator}${normalized}`;
  }

  markTruncated(): void {
    this.truncated = true;
  }

  result(
    format: SupportedDocumentFormat,
    unitType: DocumentUnitType,
    unitCount: number,
  ): ExtractedDocumentText {
    return {
      format,
      unitType,
      unitCount,
      text: this.value,
      characters: this.value.length,
      truncated: this.truncated,
    };
  }
}
