/**
 * OOXML package reading and XML-to-text utilities shared by the document
 * format parsers in document-extractor.ts: bounded archive extraction,
 * WordprocessingML / PresentationML text extraction, and spreadsheet
 * shared-string and worksheet resolution.
 */
import path from "node:path";

import { strFromU8, unzipSync, type UnzipFileInfo } from "fflate";

const MAX_ARCHIVE_ENTRIES = 1_000;
const MAX_ARCHIVE_PART_BYTES = 2 * 1_024 * 1_024;
const MAX_ARCHIVE_SELECTED_BYTES = 8 * 1_024 * 1_024;

export interface SpreadsheetSheet {
  readonly name: string;
  readonly path: string;
}


export function unzipSelectedXml(
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

export function readXmlPart(
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

export function extractWordprocessingText(xml: string): string {
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

export function extractPresentationText(xml: string): string {
  return xmlToPlainText(
    xml
      .replace(/<a:br\b[^>]*\/?>/giu, "\n")
      .replace(/<a:tab\b[^>]*\/?>/giu, "\t")
      .replace(/<\/a:p>/giu, "\n"),
  );
}

export function parseSharedStrings(xml: string): string[] {
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/giu)].map((match) =>
    extractTextElements(match[1] ?? ""),
  );
}

export function resolveSpreadsheetSheets(
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

export function extractWorksheetText(
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

export function parseXmlAttributes(value: string): Map<string, string> {
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

export function normalizeOutputText(value: string): string {
  return value
    .replace(/\r\n?/gu, "\n")
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/gu, " ").trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

export function cleanInlineText(value: string): string {
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

export function naturalCompare(left: string, right: string): number {
  return left.localeCompare(right, "en", {
    numeric: true,
    sensitivity: "base",
  });
}
