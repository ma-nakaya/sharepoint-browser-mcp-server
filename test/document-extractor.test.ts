import assert from "node:assert/strict";
import test from "node:test";

import { strToU8, zipSync } from "fflate";

import {
  extractDocumentText,
  MAX_EXTRACTED_TEXT_CHARACTERS,
} from "../src/sharepoint/document-extractor.js";

test("extracts text from a PDF page", async () => {
  const result = await extractDocumentText(createMinimalPdf("Hello PDF"), ".pdf");

  assert.equal(result.format, "pdf");
  assert.equal(result.unitType, "pages");
  assert.equal(result.unitCount, 1);
  assert.match(result.text, /\[Page 1\]/u);
  assert.match(result.text, /Hello PDF/u);
});

test("extracts DOCX body and header text while dropping deleted text", async () => {
  const archive = zipXml({
    "word/document.xml": [
      '<?xml version="1.0"?>',
      "<w:document xmlns:w=\"urn:w\">",
      "<w:body>",
      "<w:p><w:r><w:t>Hello &amp; world</w:t></w:r></w:p>",
      "<w:del><w:r><w:t>Deleted secret</w:t></w:r></w:del>",
      "<w:p><w:r><w:instrText>FIELD CODE</w:instrText></w:r>",
      "<w:r><w:t>Visible</w:t></w:r></w:p>",
      "</w:body></w:document>",
    ].join(""),
    "word/header1.xml":
      '<w:hdr xmlns:w="urn:w"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>',
  });

  const result = await extractDocumentText(archive, ".docx");

  assert.equal(result.format, "docx");
  assert.equal(result.unitCount, 2);
  assert.match(result.text, /Hello & world/u);
  assert.match(result.text, /Visible/u);
  assert.match(result.text, /Header/u);
  assert.doesNotMatch(result.text, /Deleted secret|FIELD CODE/u);
});

test("extracts XLSX sheet names, shared strings, and booleans", async () => {
  const archive = zipXml({
    "xl/workbook.xml": [
      '<workbook xmlns:r="urn:r"><sheets>',
      '<sheet name="Summary" sheetId="1" r:id="rId1"/>',
      "</sheets></workbook>",
    ].join(""),
    "xl/_rels/workbook.xml.rels":
      '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/sharedStrings.xml":
      "<sst><si><t>Hello</t></si><si><r><t>Rich</t></r><r><t> text</t></r></si></sst>",
    "xl/worksheets/sheet1.xml": [
      "<worksheet><sheetData><row r=\"1\">",
      '<c r="A1" t="s"><v>0</v></c>',
      '<c r="B1" t="s"><v>1</v></c>',
      '<c r="C1" t="b"><v>1</v></c>',
      '<c r="D1"><v>42</v></c>',
      "</row></sheetData></worksheet>",
    ].join(""),
  });

  const result = await extractDocumentText(archive, ".xlsx");

  assert.equal(result.format, "xlsx");
  assert.equal(result.unitType, "sheets");
  assert.equal(result.unitCount, 1);
  assert.match(result.text, /\[Sheet: Summary\]/u);
  assert.match(result.text, /A1: Hello/u);
  assert.match(result.text, /B1: Rich text/u);
  assert.match(result.text, /C1: TRUE/u);
  assert.match(result.text, /D1: 42/u);
});

test("extracts PPTX slide and speaker-note text", async () => {
  const archive = zipXml({
    "ppt/slides/slide1.xml":
      '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>Slide title</a:t></a:r></a:p></p:sld>',
    "ppt/notesSlides/notesSlide1.xml":
      '<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>Speaker note</a:t></a:r></a:p></p:notes>',
  });

  const result = await extractDocumentText(archive, ".pptx");

  assert.equal(result.format, "pptx");
  assert.equal(result.unitType, "slides");
  assert.equal(result.unitCount, 1);
  assert.match(result.text, /\[Slide 1\][\s\S]*Slide title/u);
  assert.match(result.text, /\[Notes 1\][\s\S]*Speaker note/u);
});

test("limits extracted text to the configured character cap", async () => {
  const archive = zipXml({
    "word/document.xml":
      `<w:document xmlns:w="urn:w"><w:body><w:p><w:t>${"x".repeat(
        MAX_EXTRACTED_TEXT_CHARACTERS + 100,
      )}</w:t></w:p></w:body></w:document>`,
  });

  const result = await extractDocumentText(archive, ".docx");

  assert.equal(result.characters, MAX_EXTRACTED_TEXT_CHARACTERS);
  assert.equal(result.truncated, true);
});

test("rejects oversized Office XML parts before decompression", async () => {
  const archive = zipXml({
    "word/document.xml":
      `<w:document><w:t>${"x".repeat(2 * 1_024 * 1_024 + 1)}</w:t></w:document>`,
  });

  await assert.rejects(
    () => extractDocumentText(archive, ".docx"),
    /XML part exceeds/u,
  );
});

test("rejects XML document type declarations", async () => {
  const archive = zipXml({
    "word/document.xml":
      '<!DOCTYPE document [<!ENTITY x "unsafe">]><w:document><w:t>&x;</w:t></w:document>',
  });

  await assert.rejects(
    () => extractDocumentText(archive, ".docx"),
    /document type declarations/u,
  );
});

test("rejects unsupported extraction types", async () => {
  await assert.rejects(
    () => extractDocumentText(new Uint8Array(), ".txt"),
    /not supported/u,
  );
});

function zipXml(files: Record<string, string>): Uint8Array {
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, value]) => [name, strToU8(value)]),
    ),
    { level: 9 },
  );
}

function createMinimalPdf(text: string): Uint8Array {
  const safeText = text.replace(/[()\\]/gu, (value) => `\\${value}`);
  const stream = `BT /F1 12 Tf 72 720 Td (${safeText}) Tj ET`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let index = 0; index < objects.length; index += 1) {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${objects[index]}\nendobj\n`;
  }
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += "0000000000 65535 f \n";
  for (const offset of offsets.slice(1)) {
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  pdf += [
    "trailer",
    `<< /Size ${objects.length + 1} /Root 1 0 R >>`,
    "startxref",
    String(xrefOffset),
    "%%EOF",
  ].join("\n");
  return new TextEncoder().encode(pdf);
}
