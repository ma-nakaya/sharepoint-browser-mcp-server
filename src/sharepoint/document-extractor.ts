import path from "node:path";

import { strFromU8, unzipSync, type UnzipFileInfo } from "fflate";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

export const MAX_EXTRACTED_TEXT_CHARACTERS = 100_000;
export const MAX_PDF_PAGES = 200;
export const MAX_DOCUMENT_OUTLINE_NODES = 500;
export const MAX_DOCUMENT_NODE_IDS = 20;
export const MAX_DOCUMENT_SEARCH_RESULTS = 20;
export const DEFAULT_DOCUMENT_SEARCH_RESULTS = 10;
export const MAX_DOCUMENT_SEARCH_QUERY_CHARACTERS = 200;

const MAX_ARCHIVE_ENTRIES = 1_000;
const MAX_ARCHIVE_PART_BYTES = 2 * 1_024 * 1_024;
const MAX_ARCHIVE_SELECTED_BYTES = 8 * 1_024 * 1_024;
const MAX_OFFICE_UNITS = 200;
const MAX_SPREADSHEET_CELLS = 20_000;

export type SupportedDocumentFormat = "pdf" | "docx" | "xlsx" | "pptx";
export type DocumentUnitType = "pages" | "parts" | "sheets" | "slides";
export type DocumentNodeKind = "page" | "section" | "part" | "sheet" | "slide";
export type DocumentNodePositionType =
  | "page"
  | "paragraph"
  | "part"
  | "sheet"
  | "slide";

export interface ExtractedDocumentText {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly text: string;
  readonly characters: number;
  readonly truncated: boolean;
}

export interface DocumentOutlineNode {
  readonly nodeId: string;
  readonly title: string;
  readonly kind: DocumentNodeKind;
  readonly positionType: DocumentNodePositionType;
  readonly startIndex: number;
  readonly endIndex: number;
  readonly locator: string;
  readonly characters: number;
  readonly truncated: boolean;
  readonly preview?: string;
  readonly children: readonly DocumentOutlineNode[];
}

export interface ExtractedDocumentOutline {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly nodeCount: number;
  readonly truncated: boolean;
  readonly nodes: readonly DocumentOutlineNode[];
}

export interface ExtractedDocumentNode {
  readonly nodeId: string;
  readonly title: string;
  readonly kind: DocumentNodeKind;
  readonly positionType: DocumentNodePositionType;
  readonly startIndex: number;
  readonly endIndex: number;
  readonly locator: string;
  readonly text: string;
  readonly characters: number;
  readonly truncated: boolean;
}

export interface ExtractedDocumentNodes {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly requestedNodeIds: readonly string[];
  readonly returnedNodes: number;
  readonly characters: number;
  readonly truncated: boolean;
  readonly nodes: readonly ExtractedDocumentNode[];
}

export interface DocumentSearchMatch {
  readonly nodeId: string;
  readonly title: string;
  readonly kind: DocumentNodeKind;
  readonly positionType: DocumentNodePositionType;
  readonly startIndex: number;
  readonly endIndex: number;
  readonly locator: string;
  readonly score: number;
  readonly snippet: string;
}

export interface ExtractedDocumentSearch {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly query: string;
  readonly matchedNodes: number;
  readonly returnedNodes: number;
  readonly truncated: boolean;
  readonly results: readonly DocumentSearchMatch[];
}

interface SpreadsheetSheet {
  readonly name: string;
  readonly path: string;
}

interface InternalDocumentNode {
  readonly nodeId: string;
  readonly title: string;
  readonly kind: DocumentNodeKind;
  readonly positionType: DocumentNodePositionType;
  readonly startIndex: number;
  readonly endIndex: number;
  readonly locator: string;
  readonly text: string;
  readonly searchText: string;
  readonly truncated: boolean;
  readonly children: readonly InternalDocumentNode[];
}

interface ParsedDocumentStructure {
  readonly format: SupportedDocumentFormat;
  readonly unitType: DocumentUnitType;
  readonly unitCount: number;
  readonly truncated: boolean;
  readonly nodes: readonly InternalDocumentNode[];
}

interface WordParagraph {
  readonly text: string;
  readonly headingLevel?: number;
}

interface WordSectionDraft {
  readonly level: number;
  readonly node: MutableInternalDocumentNode;
}

interface MutableInternalDocumentNode {
  nodeId: string;
  title: string;
  kind: DocumentNodeKind;
  positionType: DocumentNodePositionType;
  startIndex: number;
  endIndex: number;
  locator: string;
  text: string;
  searchText: string;
  truncated: boolean;
  children: MutableInternalDocumentNode[];
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

export async function extractDocumentOutline(
  data: Uint8Array,
  extension: string,
): Promise<ExtractedDocumentOutline> {
  const parsed = await parseDocumentStructure(data, extension);
  const flattened = flattenInternalNodes(parsed.nodes);
  return {
    format: parsed.format,
    unitType: parsed.unitType,
    unitCount: parsed.unitCount,
    nodeCount: flattened.length,
    truncated: parsed.truncated,
    nodes: parsed.nodes.map(toOutlineNode),
  };
}

export async function extractDocumentNodes(
  data: Uint8Array,
  extension: string,
  nodeIds: readonly string[],
): Promise<ExtractedDocumentNodes> {
  const requestedNodeIds = normalizeRequestedNodeIds(nodeIds);
  const parsed = await parseDocumentStructure(data, extension);
  const availableNodes = new Map(
    flattenInternalNodes(parsed.nodes).map((node) => [node.nodeId, node]),
  );
  const missingNodeIds = requestedNodeIds.filter(
    (nodeId) => !availableNodes.has(nodeId),
  );
  if (missingNodeIds.length > 0) {
    throw new Error(`Unknown document node ID: ${missingNodeIds[0]}.`);
  }

  const nodes: ExtractedDocumentNode[] = [];
  let characters = 0;
  let truncated = parsed.truncated;
  for (const nodeId of requestedNodeIds) {
    const node = availableNodes.get(nodeId);
    if (!node) {
      continue;
    }
    const available = MAX_EXTRACTED_TEXT_CHARACTERS - characters;
    if (available <= 0) {
      truncated = true;
      break;
    }
    const text = node.text.slice(0, available);
    const nodeTruncated = node.truncated || text.length < node.text.length;
    nodes.push({
      nodeId: node.nodeId,
      title: node.title,
      kind: node.kind,
      positionType: node.positionType,
      startIndex: node.startIndex,
      endIndex: node.endIndex,
      locator: node.locator,
      text,
      characters: text.length,
      truncated: nodeTruncated,
    });
    characters += text.length;
    truncated ||= nodeTruncated;
  }

  return {
    format: parsed.format,
    unitType: parsed.unitType,
    unitCount: parsed.unitCount,
    requestedNodeIds,
    returnedNodes: nodes.length,
    characters,
    truncated,
    nodes,
  };
}

export async function searchDocumentStructure(
  data: Uint8Array,
  extension: string,
  query: string,
  maxResults = DEFAULT_DOCUMENT_SEARCH_RESULTS,
): Promise<ExtractedDocumentSearch> {
  const normalizedQuery = normalizeDocumentSearchQuery(query);
  const normalizedLimit = normalizeDocumentSearchResultLimit(maxResults);
  const parsed = await parseDocumentStructure(data, extension);
  const queryLower = normalizedQuery.toLowerCase();
  const tokens = [...new Set(
    queryLower.split(/\s+/u).filter(Boolean),
  )];
  const matches = flattenInternalNodes(parsed.nodes)
    .map((node) => scoreDocumentNode(node, queryLower, tokens))
    .filter((match): match is DocumentSearchMatch => match !== undefined)
    .sort((left, right) =>
      right.score - left.score ||
      left.endIndex - left.startIndex - (right.endIndex - right.startIndex) ||
      left.startIndex - right.startIndex ||
      left.nodeId.localeCompare(right.nodeId, "en"),
    );

  return {
    format: parsed.format,
    unitType: parsed.unitType,
    unitCount: parsed.unitCount,
    query: normalizedQuery,
    matchedNodes: matches.length,
    returnedNodes: Math.min(matches.length, normalizedLimit),
    truncated: parsed.truncated || matches.length > normalizedLimit,
    results: matches.slice(0, normalizedLimit),
  };
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

async function parseDocumentStructure(
  data: Uint8Array,
  extension: string,
): Promise<ParsedDocumentStructure> {
  switch (extension.toLowerCase()) {
    case ".pdf":
      return parsePdfStructure(data);
    case ".docx":
      return parseDocxStructure(data);
    case ".xlsx":
      return parseXlsxStructure(data);
    case ".pptx":
      return parsePptxStructure(data);
    default:
      throw new Error("Document type is not supported for structured extraction.");
  }
}

async function parsePdfStructure(
  data: Uint8Array,
): Promise<ParsedDocumentStructure> {
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

    const nodes: InternalDocumentNode[] = [];
    let truncated = false;
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await document.getPage(pageNumber);
      const content = await page.getTextContent();
      let rawText = "";
      for (const item of content.items) {
        if (!("str" in item)) {
          continue;
        }
        rawText += item.str;
        rawText += item.hasEOL ? "\n" : " ";
      }
      page.cleanup();
      const limited = limitNodeText(rawText);
      truncated ||= limited.truncated;
      nodes.push({
        nodeId: createNodeId("page", pageNumber),
        title: createPositionTitle("Page", pageNumber, limited.text),
        kind: "page",
        positionType: "page",
        startIndex: pageNumber,
        endIndex: pageNumber,
        locator: `page:${pageNumber}`,
        text: limited.text,
        searchText: limited.text,
        truncated: limited.truncated,
        children: [],
      });
    }
    return {
      format: "pdf",
      unitType: "pages",
      unitCount: document.numPages,
      truncated,
      nodes,
    };
  } catch (error) {
    if (error instanceof Error && /exceeds the \d+-page/u.test(error.message)) {
      throw error;
    }
    throw new Error("PDF structure could not be extracted.");
  } finally {
    if (document) {
      await document.cleanup().catch(() => undefined);
    }
    await loadingTask.destroy().catch(() => undefined);
  }
}

function parseDocxStructure(data: Uint8Array): ParsedDocumentStructure {
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

  const bodyXml = readXmlPart(files, "word/document.xml");
  const paragraphs = extractWordParagraphs(bodyXml);
  const sectionResult = buildWordSectionNodes(paragraphs);
  const nodes: MutableInternalDocumentNode[] = [...sectionResult.nodes];
  let truncated = sectionResult.truncated;
  let partIndex = 0;

  for (const partName of orderedParts) {
    if (partName === "word/document.xml") {
      continue;
    }
    partIndex += 1;
    const label = partName
      .replace(/^word\//u, "")
      .replace(/\.xml$/u, "");
    const limited = limitNodeText(
      extractWordprocessingText(readXmlPart(files, partName)),
    );
    if (countMutableNodes(nodes) >= MAX_DOCUMENT_OUTLINE_NODES) {
      truncated = true;
      break;
    }
    nodes.push({
      nodeId: createNodeId("part", partIndex),
      title: formatNodeTitle(label),
      kind: "part",
      positionType: "part",
      startIndex: partIndex,
      endIndex: partIndex,
      locator: `part:${label}`,
      text: limited.text,
      searchText: limited.text,
      truncated: limited.truncated,
      children: [],
    });
    truncated ||= limited.truncated;
  }

  return {
    format: "docx",
    unitType: "parts",
    unitCount: orderedParts.length,
    truncated,
    nodes,
  };
}

function parseXlsxStructure(data: Uint8Array): ParsedDocumentStructure {
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
  const nodes: InternalDocumentNode[] = [];
  let cellsRemaining = MAX_SPREADSHEET_CELLS;
  let truncated = false;

  for (let index = 0; index < sheets.length; index += 1) {
    const sheet = sheets[index];
    if (!sheet) {
      continue;
    }
    let text = "";
    let nodeTruncated = false;
    if (cellsRemaining > 0) {
      const result = extractWorksheetText(
        readXmlPart(files, sheet.path),
        sharedStrings,
        cellsRemaining,
      );
      cellsRemaining -= result.cells;
      text = result.text;
      nodeTruncated = result.limitReached;
    } else {
      nodeTruncated = true;
    }
    const limited = limitNodeText(text);
    nodeTruncated ||= limited.truncated;
    truncated ||= nodeTruncated;
    const sheetNumber = index + 1;
    nodes.push({
      nodeId: createNodeId("sheet", sheetNumber),
      title: formatNodeTitle(sheet.name),
      kind: "sheet",
      positionType: "sheet",
      startIndex: sheetNumber,
      endIndex: sheetNumber,
      locator: `sheet:${sheet.name}`,
      text: limited.text,
      searchText: limited.text,
      truncated: nodeTruncated,
      children: [],
    });
  }

  return {
    format: "xlsx",
    unitType: "sheets",
    unitCount: sheets.length,
    truncated,
    nodes,
  };
}

function parsePptxStructure(data: Uint8Array): ParsedDocumentStructure {
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

  const nodes: InternalDocumentNode[] = [];
  let truncated = false;
  for (let index = 0; index < slidePaths.length; index += 1) {
    const slidePath = slidePaths[index];
    if (!slidePath) {
      continue;
    }
    const slideNumber = index + 1;
    const slideText = extractPresentationText(readXmlPart(files, slidePath));
    const notesPath = `ppt/notesSlides/notesSlide${slideNumber}.xml`;
    const notesText = files[notesPath]
      ? extractPresentationText(readXmlPart(files, notesPath))
      : "";
    const combined = notesText
      ? `${slideText}\n\nNotes:\n${notesText}`
      : slideText;
    const limited = limitNodeText(combined);
    truncated ||= limited.truncated;
    nodes.push({
      nodeId: createNodeId("slide", slideNumber),
      title: createPositionTitle("Slide", slideNumber, slideText),
      kind: "slide",
      positionType: "slide",
      startIndex: slideNumber,
      endIndex: slideNumber,
      locator: `slide:${slideNumber}`,
      text: limited.text,
      searchText: limited.text,
      truncated: limited.truncated,
      children: [],
    });
  }

  return {
    format: "pptx",
    unitType: "slides",
    unitCount: slidePaths.length,
    truncated,
    nodes,
  };
}

function extractWordParagraphs(xml: string): WordParagraph[] {
  const paragraphs: WordParagraph[] = [];
  for (const match of xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/giu)) {
    const body = match[1] ?? "";
    const text = extractWordprocessingText(match[0]);
    const propertyXml = /<w:pPr\b[^>]*>([\s\S]*?)<\/w:pPr>/iu.exec(body)?.[1];
    const headingLevel = propertyXml
      ? extractWordHeadingLevel(propertyXml)
      : undefined;
    paragraphs.push({
      text,
      ...(headingLevel !== undefined ? { headingLevel } : {}),
    });
  }
  if (paragraphs.length === 0) {
    const text = extractWordprocessingText(xml);
    return text ? [{ text }] : [];
  }
  return paragraphs;
}

function extractWordHeadingLevel(propertyXml: string): number | undefined {
  const outlineMatch = /<w:outlineLvl\b([^>]*)\/?>/iu.exec(propertyXml);
  if (outlineMatch) {
    const value = getXmlAttribute(outlineMatch[1] ?? "", "w:val", "val");
    const level = value === undefined ? Number.NaN : Number(value);
    if (Number.isInteger(level) && level >= 0 && level <= 8) {
      return level + 1;
    }
  }

  const styleMatch = /<w:pStyle\b([^>]*)\/?>/iu.exec(propertyXml);
  if (!styleMatch) {
    return undefined;
  }
  const style = getXmlAttribute(styleMatch[1] ?? "", "w:val", "val");
  const match = style?.match(/^(?:Heading|見出し)\s*([1-9])$/iu);
  return match?.[1] ? Number(match[1]) : undefined;
}

function buildWordSectionNodes(
  paragraphs: readonly WordParagraph[],
): { readonly nodes: MutableInternalDocumentNode[]; readonly truncated: boolean } {
  const headingIndexes = paragraphs
    .map((paragraph, index) =>
      paragraph.headingLevel !== undefined && paragraph.text
        ? index
        : undefined,
    )
    .filter((index): index is number => index !== undefined);
  const roots: MutableInternalDocumentNode[] = [];
  const drafts: WordSectionDraft[] = [];
  let sectionNumber = 0;
  let truncated = false;

  const firstHeading = headingIndexes[0];
  if (firstHeading === undefined) {
    const text = paragraphs.map((paragraph) => paragraph.text).filter(Boolean).join("\n");
    const limited = limitNodeText(text);
    roots.push({
      nodeId: createNodeId("section", 1),
      title: createPositionTitle("Document", undefined, text),
      kind: "section",
      positionType: "paragraph",
      startIndex: paragraphs.length > 0 ? 1 : 0,
      endIndex: paragraphs.length,
      locator: paragraphs.length > 0
        ? `paragraphs:1-${paragraphs.length}`
        : "paragraphs:0-0",
      text: limited.text,
      searchText: limited.text,
      truncated: limited.truncated,
      children: [],
    });
    return { nodes: roots, truncated: limited.truncated };
  }

  const introductionText = paragraphs
    .slice(0, firstHeading)
    .map((paragraph) => paragraph.text)
    .filter(Boolean)
    .join("\n");
  if (introductionText) {
    sectionNumber += 1;
    const limited = limitNodeText(introductionText);
    roots.push({
      nodeId: createNodeId("section", sectionNumber),
      title: createPositionTitle("Introduction", undefined, introductionText),
      kind: "section",
      positionType: "paragraph",
      startIndex: 1,
      endIndex: firstHeading,
      locator: `paragraphs:1-${firstHeading}`,
      text: limited.text,
      searchText: limited.text,
      truncated: limited.truncated,
      children: [],
    });
    truncated ||= limited.truncated;
  }

  for (let headingOffset = 0; headingOffset < headingIndexes.length; headingOffset += 1) {
    if (sectionNumber >= MAX_DOCUMENT_OUTLINE_NODES) {
      truncated = true;
      break;
    }
    const paragraphIndex = headingIndexes[headingOffset];
    if (paragraphIndex === undefined) {
      continue;
    }
    const paragraph = paragraphs[paragraphIndex];
    if (!paragraph?.headingLevel) {
      continue;
    }
    const nextHeadingIndex = headingIndexes[headingOffset + 1] ?? paragraphs.length;
    let sectionEnd = paragraphs.length;
    for (
      let followingOffset = headingOffset + 1;
      followingOffset < headingIndexes.length;
      followingOffset += 1
    ) {
      const followingIndex = headingIndexes[followingOffset];
      const following = followingIndex === undefined
        ? undefined
        : paragraphs[followingIndex];
      if (
        followingIndex !== undefined &&
        following?.headingLevel !== undefined &&
        following.headingLevel <= paragraph.headingLevel
      ) {
        sectionEnd = followingIndex;
        break;
      }
    }
    const text = paragraphs
      .slice(paragraphIndex, sectionEnd)
      .map((value) => value.text)
      .filter(Boolean)
      .join("\n");
    const searchText = paragraphs
      .slice(paragraphIndex, nextHeadingIndex)
      .map((value) => value.text)
      .filter(Boolean)
      .join("\n");
    const limitedText = limitNodeText(text);
    const limitedSearchText = limitNodeText(searchText);
    sectionNumber += 1;
    drafts.push({
      level: paragraph.headingLevel,
      node: {
        nodeId: createNodeId("section", sectionNumber),
        title: formatNodeTitle(paragraph.text),
        kind: "section",
        positionType: "paragraph",
        startIndex: paragraphIndex + 1,
        endIndex: sectionEnd,
        locator: `paragraphs:${paragraphIndex + 1}-${sectionEnd}`,
        text: limitedText.text,
        searchText: limitedSearchText.text,
        truncated: limitedText.truncated || limitedSearchText.truncated,
        children: [],
      },
    });
    truncated ||= limitedText.truncated || limitedSearchText.truncated;
  }

  const stack: WordSectionDraft[] = [];
  for (const draft of drafts) {
    while (stack.length > 0 && (stack.at(-1)?.level ?? 0) >= draft.level) {
      stack.pop();
    }
    const parent = stack.at(-1);
    if (parent) {
      parent.node.children.push(draft.node);
    } else {
      roots.push(draft.node);
    }
    stack.push(draft);
  }
  return { nodes: roots, truncated };
}

function normalizeRequestedNodeIds(nodeIds: readonly string[]): string[] {
  if (nodeIds.length === 0 || nodeIds.length > MAX_DOCUMENT_NODE_IDS) {
    throw new Error(
      `Document node selection must contain between 1 and ${MAX_DOCUMENT_NODE_IDS} IDs.`,
    );
  }
  const normalized = [...new Set(nodeIds.map((nodeId) => nodeId.trim()))];
  if (
    normalized.some(
      (nodeId) => !/^(?:page|section|part|sheet|slide)-\d{4}$/u.test(nodeId),
    )
  ) {
    throw new Error("Document node IDs must use the server-generated node format.");
  }
  return normalized;
}

function normalizeDocumentSearchQuery(query: string): string {
  const normalized = query.trim();
  if (normalized.length === 0) {
    throw new Error("Document search query must not be empty.");
  }
  if (normalized.length > MAX_DOCUMENT_SEARCH_QUERY_CHARACTERS) {
    throw new Error(
      `Document search query must be ${MAX_DOCUMENT_SEARCH_QUERY_CHARACTERS} characters or fewer.`,
    );
  }
  return normalized;
}

function normalizeDocumentSearchResultLimit(value: number): number {
  if (
    !Number.isSafeInteger(value) ||
    value < 1 ||
    value > MAX_DOCUMENT_SEARCH_RESULTS
  ) {
    throw new Error(
      `Document search result limit must be between 1 and ${MAX_DOCUMENT_SEARCH_RESULTS}.`,
    );
  }
  return value;
}

function scoreDocumentNode(
  node: InternalDocumentNode,
  query: string,
  tokens: readonly string[],
): DocumentSearchMatch | undefined {
  const title = node.title.toLowerCase();
  const text = node.searchText.toLowerCase();
  const titlePhraseMatches = countOccurrences(title, query, 3);
  const textPhraseMatches = countOccurrences(text, query, 10);
  const tokenMatches = tokens.map((token) => ({
    token,
    title: countOccurrences(title, token, 3),
    text: countOccurrences(text, token, 10),
  }));
  if (
    titlePhraseMatches === 0 &&
    textPhraseMatches === 0 &&
    tokenMatches.every((match) => match.title === 0 && match.text === 0)
  ) {
    return undefined;
  }

  const allTokensMatched = tokenMatches.every(
    (match) => match.title > 0 || match.text > 0,
  );
  const score = Math.round(
    titlePhraseMatches * 100 +
    textPhraseMatches * 30 +
    tokenMatches.reduce(
      (total, match) => total + match.title * 20 + match.text * 5,
      0,
    ) +
    (allTokensMatched ? 15 : 0),
  );
  const firstNeedle = text.includes(query)
    ? query
    : tokenMatches.find((match) => match.text > 0)?.token;
  return {
    nodeId: node.nodeId,
    title: node.title,
    kind: node.kind,
    positionType: node.positionType,
    startIndex: node.startIndex,
    endIndex: node.endIndex,
    locator: node.locator,
    score,
    snippet: createSearchSnippet(node.searchText, firstNeedle),
  };
}

function countOccurrences(value: string, needle: string, limit: number): number {
  if (!needle) {
    return 0;
  }
  let count = 0;
  let offset = 0;
  while (count < limit) {
    const index = value.indexOf(needle, offset);
    if (index < 0) {
      break;
    }
    count += 1;
    offset = index + Math.max(needle.length, 1);
  }
  return count;
}

function createSearchSnippet(value: string, needle: string | undefined): string {
  const normalized = cleanInlineText(value);
  if (!normalized) {
    return "";
  }
  const lower = normalized.toLowerCase();
  const matchIndex = needle ? lower.indexOf(needle) : -1;
  const start = Math.max(0, matchIndex >= 0 ? matchIndex - 100 : 0);
  const end = Math.min(normalized.length, start + 320);
  return `${start > 0 ? "…" : ""}${normalized.slice(start, end)}${
    end < normalized.length ? "…" : ""
  }`;
}

function flattenInternalNodes(
  nodes: readonly InternalDocumentNode[],
): InternalDocumentNode[] {
  const flattened: InternalDocumentNode[] = [];
  const visit = (node: InternalDocumentNode): void => {
    flattened.push(node);
    for (const child of node.children) {
      visit(child);
    }
  };
  for (const node of nodes) {
    visit(node);
  }
  return flattened;
}

function toOutlineNode(node: InternalDocumentNode): DocumentOutlineNode {
  const preview = createSearchSnippet(node.searchText || node.text, undefined)
    .slice(0, 240);
  return {
    nodeId: node.nodeId,
    title: node.title,
    kind: node.kind,
    positionType: node.positionType,
    startIndex: node.startIndex,
    endIndex: node.endIndex,
    locator: node.locator,
    characters: node.text.length,
    truncated: node.truncated,
    ...(preview ? { preview } : {}),
    children: node.children.map(toOutlineNode),
  };
}

function limitNodeText(
  value: string,
): { readonly text: string; readonly truncated: boolean } {
  const normalized = normalizeOutputText(value);
  return normalized.length > MAX_EXTRACTED_TEXT_CHARACTERS
    ? {
        text: normalized.slice(0, MAX_EXTRACTED_TEXT_CHARACTERS),
        truncated: true,
      }
    : { text: normalized, truncated: false };
}

function createNodeId(
  prefix: "page" | "section" | "part" | "sheet" | "slide",
  index: number,
): string {
  return `${prefix}-${String(index).padStart(4, "0")}`;
}

function createPositionTitle(
  label: string,
  index: number | undefined,
  text: string,
): string {
  const firstLine = normalizeOutputText(text).split("\n").find(Boolean);
  const position = index === undefined ? label : `${label} ${index}`;
  return firstLine
    ? `${position} — ${formatNodeTitle(firstLine, 120)}`
    : position;
}

function formatNodeTitle(value: string, maxCharacters = 160): string {
  const normalized = cleanInlineText(value);
  return (normalized || "Untitled").slice(0, maxCharacters);
}

function getXmlAttribute(
  value: string,
  ...names: readonly string[]
): string | undefined {
  const attributes = parseXmlAttributes(value);
  return names.map((name) => attributes.get(name)).find(
    (attribute): attribute is string => attribute !== undefined,
  );
}

function countMutableNodes(nodes: readonly MutableInternalDocumentNode[]): number {
  return nodes.reduce(
    (count, node) => count + 1 + countMutableNodes(node.children),
    0,
  );
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
