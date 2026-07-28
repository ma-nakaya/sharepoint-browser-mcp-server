import assert from "node:assert/strict";
import test from "node:test";

import {
  buildFileDownloadApiPath,
  buildFolderChildrenApiPaths,
  normalizeSharePointFileUrl,
  normalizeSharePointFolderUrl,
  parseDocumentLibrariesResponse,
  parseFileEntriesResponse,
  parseFileMetadataResponse,
  parseFolderEntriesResponse,
} from "../src/sharepoint/file-content.js";
import { buildSharePointApiUrl } from "../src/sharepoint/url-guard.js";

const SITE_URL = "https://example.sharepoint.com/teams/genai";

test("normalizes site files and folders and rejects unsafe boundaries", () => {
  const folder = normalizeSharePointFolderUrl(
    SITE_URL,
    "/teams/genai/Shared%20Documents/Plans",
  );
  const file = normalizeSharePointFileUrl(
    SITE_URL,
    "https://example.sharepoint.com/teams/genai/Shared%20Documents/Plan.pdf?download=1",
  );

  assert.equal(
    folder.serverRelativeUrl,
    "/teams/genai/Shared Documents/Plans",
  );
  assert.equal(
    file.url,
    "https://example.sharepoint.com/teams/genai/Shared%20Documents/Plan.pdf",
  );
  assert.throws(
    () =>
      normalizeSharePointFolderUrl(
        SITE_URL,
        "/teams/genai-other/Shared Documents",
      ),
    /configured SharePoint site/u,
  );
  assert.throws(
    () =>
      normalizeSharePointFileUrl(
        SITE_URL,
        "/teams/genai/Shared Documents/run.exe",
      ),
    /not allowed/u,
  );
});

test("builds ResourcePath endpoints without URL fragments", () => {
  const filePath = "/teams/genai/Shared Documents/100% # plan.pdf";
  const downloadPath = buildFileDownloadApiPath(filePath);
  const childrenPaths = buildFolderChildrenApiPaths(
    "/teams/genai/Shared Documents",
    20,
  );
  const url = buildSharePointApiUrl(SITE_URL, downloadPath);

  assert.match(downloadPath, /100%25%20%23%20plan\.pdf/u);
  assert.match(downloadPath, /\/\$value$/u);
  assert.equal(url.hash, "");
  assert.match(childrenPaths.files, /\$top=20/u);
  assert.match(childrenPaths.folders, /\/Folders/u);
});

test("parses visible document libraries and drops out-of-site roots", () => {
  const body = JSON.stringify({
    value: [
      {
        Title: "Documents",
        ItemCount: 12,
        RootFolder: {
          ServerRelativeUrl: "/teams/genai/Shared Documents",
        },
      },
      {
        Title: "Other",
        ItemCount: 1,
        RootFolder: {
          ServerRelativeUrl: "/teams/other/Shared Documents",
        },
      },
    ],
  });

  const libraries = parseDocumentLibrariesResponse(body, SITE_URL);

  assert.equal(libraries.length, 1);
  assert.equal(libraries[0]?.title, "Documents");
  assert.equal(libraries[0]?.itemCount, 12);
});

test("parses only direct child folders and files", () => {
  const folder = normalizeSharePointFolderUrl(
    SITE_URL,
    "/teams/genai/Shared Documents",
  );
  const foldersBody = JSON.stringify({
    value: [
      {
        Name: "Plans",
        ServerRelativeUrl: "/teams/genai/Shared Documents/Plans",
        ItemCount: 3,
      },
      {
        Name: "Nested",
        ServerRelativeUrl: "/teams/genai/Shared Documents/Plans/Nested",
        ItemCount: 1,
      },
    ],
  });
  const filesBody = JSON.stringify({
    value: [
      {
        Name: "Plan.pdf",
        ServerRelativeUrl: "/teams/genai/Shared Documents/Plan.pdf",
        Length: "123",
        TimeLastModified: "2026-07-28T00:00:00Z",
        MajorVersion: 4,
      },
      {
        Name: "run.exe",
        ServerRelativeUrl: "/teams/genai/Shared Documents/run.exe",
        Length: "20",
      },
    ],
  });

  const folders = parseFolderEntriesResponse(
    foldersBody,
    SITE_URL,
    folder,
  );
  const files = parseFileEntriesResponse(filesBody, SITE_URL, folder);

  assert.equal(folders.length, 1);
  assert.equal(folders[0]?.name, "Plans");
  assert.equal(files.length, 2);
  assert.equal(files[0]?.downloadable, true);
  assert.equal(files[1]?.downloadable, false);
});

test("validates downloaded file metadata against the requested path", () => {
  const file = normalizeSharePointFileUrl(
    SITE_URL,
    "/teams/genai/Shared Documents/Plan.pdf",
  );
  const metadata = parseFileMetadataResponse(
    JSON.stringify({
      Name: "Plan.pdf",
      ServerRelativeUrl: file.serverRelativeUrl,
      Length: "4",
      MajorVersion: 1,
    }),
    SITE_URL,
    file,
  );

  assert.equal(metadata.sizeBytes, 4);
  assert.equal(metadata.mimeType, "application/pdf");
  assert.throws(
    () =>
      parseFileMetadataResponse(
        JSON.stringify({
          Name: "Other.pdf",
          ServerRelativeUrl: "/teams/genai/Shared Documents/Other.pdf",
          Length: "4",
        }),
        SITE_URL,
        file,
      ),
    /unexpected file/u,
  );
});
