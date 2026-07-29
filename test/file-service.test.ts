import assert from "node:assert/strict";
import test from "node:test";

import {
  MAX_DOCUMENT_SOURCE_BYTES,
  MAX_DOWNLOAD_BYTES,
} from "../src/sharepoint/file-content.js";
import { SharePointFileService } from "../src/sharepoint/file-service.js";
import type {
  SharePointBinaryResponse,
  SharePointBinaryTransport,
  SharePointResponse,
} from "../src/sharepoint/http.js";

const SITE_URL = "https://example.sharepoint.com/sites/genai";
const FILE_URL = "/sites/genai/Shared Documents/readme.txt";

type JsonHandler = (apiPath: string) => SharePointResponse | Error;
type BinaryHandler = (apiPath: string) => SharePointBinaryResponse | Error;

class FakeBinaryTransport implements SharePointBinaryTransport {
  public fallbackJsonCalls = 0;
  public primaryBinaryCalls = 0;
  public fallbackBinaryCalls = 0;
  public primaryBinaryLimits: Array<number | undefined> = [];
  public fallbackBinaryLimits: Array<number | undefined> = [];

  constructor(
    private readonly primaryJson: JsonHandler,
    private readonly fallbackJson: JsonHandler,
    private readonly primaryBinary: BinaryHandler,
    private readonly fallbackBinary: BinaryHandler,
  ) {}

  async get(apiPath: string): Promise<SharePointResponse> {
    return resolve(this.primaryJson(apiPath));
  }

  async getViaPage(apiPath: string): Promise<SharePointResponse> {
    this.fallbackJsonCalls += 1;
    return resolve(this.fallbackJson(apiPath));
  }

  async getBinary(
    apiPath: string,
    maxBytes?: number,
  ): Promise<SharePointBinaryResponse> {
    this.primaryBinaryCalls += 1;
    this.primaryBinaryLimits.push(maxBytes);
    return resolve(this.primaryBinary(apiPath));
  }

  async getBinaryViaPage(
    apiPath: string,
    maxBytes?: number,
  ): Promise<SharePointBinaryResponse> {
    this.fallbackBinaryCalls += 1;
    this.fallbackBinaryLimits.push(maxBytes);
    return resolve(this.fallbackBinary(apiPath));
  }

  async close(): Promise<void> {}
}

function jsonResponse(
  body: string,
  method: "api-request" | "browser-fetch" = "api-request",
): SharePointResponse {
  return {
    status: 200,
    contentType: "application/json;odata=nometadata",
    body,
    bodyTruncated: false,
    method,
  };
}

function jsonStatusResponse(
  status: number,
  method: "api-request" | "browser-fetch" = "api-request",
): SharePointResponse {
  return {
    status,
    contentType: "application/json;odata=nometadata",
    body: "",
    bodyTruncated: false,
    method,
  };
}

function binaryResponse(
  body: Uint8Array,
  contentType = "application/octet-stream",
  method: "api-request" | "browser-fetch" = "api-request",
): SharePointBinaryResponse {
  return {
    status: 200,
    contentType,
    body,
    bodyTruncated: false,
    method,
  };
}

function binaryStatusResponse(
  status: number,
  method: "api-request" | "browser-fetch" = "api-request",
): SharePointBinaryResponse {
  return {
    status,
    contentType: "application/octet-stream",
    body: new Uint8Array(),
    bodyTruncated: false,
    method,
  };
}

function fileMetadata(size: number): string {
  return JSON.stringify({
    Name: "readme.txt",
    ServerRelativeUrl: FILE_URL,
    Length: String(size),
    TimeLastModified: "2026-07-28T00:00:00Z",
  });
}

test("lists document libraries from the primary request", async () => {
  const transport = new FakeBinaryTransport(
    () =>
      jsonResponse(
        JSON.stringify({
          value: [
            {
              Title: "Documents",
              ItemCount: 2,
              RootFolder: {
                ServerRelativeUrl: "/sites/genai/Shared Documents",
              },
            },
          ],
        }),
      ),
    () => new Error("fallback must not run"),
    () => new Error("binary must not run"),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.listDocumentLibraries();

  assert.equal(result.returnedLibraries, 1);
  assert.equal(result.libraries[0]?.title, "Documents");
  assert.equal(transport.fallbackJsonCalls, 0);
});

test("lists document libraries through browser fetch after a primary 403", async () => {
  const librariesBody = JSON.stringify({
    value: [
      {
        Title: "Documents",
        ItemCount: 2,
        RootFolder: {
          ServerRelativeUrl: "/sites/genai/Shared Documents",
        },
      },
    ],
  });
  const transport = new FakeBinaryTransport(
    () => jsonStatusResponse(403),
    () => jsonResponse(librariesBody, "browser-fetch"),
    () => new Error("binary must not run"),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.listDocumentLibraries();

  assert.equal(result.returnedLibraries, 1);
  assert.equal(result.method, "browser-fetch");
  assert.equal(transport.fallbackJsonCalls, 1);
});

test("reports access denied when both library request methods return 403", async () => {
  const transport = new FakeBinaryTransport(
    () => jsonStatusResponse(403),
    () => jsonStatusResponse(403, "browser-fetch"),
    () => new Error("binary must not run"),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  await assert.rejects(
    () => service.listDocumentLibraries(),
    /cannot read the requested SharePoint document libraries/u,
  );
  assert.equal(transport.fallbackJsonCalls, 1);
});

test("lists direct folder children", async () => {
  const transport = new FakeBinaryTransport(
    (apiPath) =>
      apiPath.includes("/Folders?")
        ? jsonResponse(
            JSON.stringify({
              value: [
                {
                  Name: "Plans",
                  ServerRelativeUrl: "/sites/genai/Shared Documents/Plans",
                  ItemCount: 1,
                },
              ],
            }),
          )
        : jsonResponse(
            JSON.stringify({
              value: [
                {
                  Name: "readme.txt",
                  ServerRelativeUrl: FILE_URL,
                  Length: "5",
                },
              ],
            }),
          ),
    () => new Error("fallback must not run"),
    () => new Error("binary must not run"),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.listFolder(
    "/sites/genai/Shared Documents",
    10,
  );

  assert.equal(result.returnedFolders, 1);
  assert.equal(result.returnedFiles, 1);
  assert.equal(result.files[0]?.downloadable, true);
});

test("downloads a file and returns integrity metadata", async () => {
  const bytes = new TextEncoder().encode("hello");
  const transport = new FakeBinaryTransport(
    () => jsonResponse(fileMetadata(bytes.byteLength)),
    () => new Error("fallback must not run"),
    () => binaryResponse(bytes),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.downloadFile(FILE_URL);

  assert.deepEqual(result.data, bytes);
  assert.equal(result.mimeType, "text/plain");
  assert.equal(
    result.sha256,
    "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824",
  );
});

test("uses browser fetch when the primary binary response is HTML", async () => {
  const bytes = new TextEncoder().encode("hello");
  const transport = new FakeBinaryTransport(
    () => jsonResponse(fileMetadata(bytes.byteLength)),
    () => new Error("JSON fallback must not run"),
    () => binaryResponse(new TextEncoder().encode("<html>"), "text/html"),
    () => binaryResponse(bytes, "text/plain", "browser-fetch"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.downloadFile(FILE_URL);

  assert.equal(result.method, "browser-fetch");
  assert.equal(transport.fallbackBinaryCalls, 1);
});

test("uses browser fetch when the primary binary request returns 403", async () => {
  const bytes = new TextEncoder().encode("hello");
  const transport = new FakeBinaryTransport(
    () => jsonResponse(fileMetadata(bytes.byteLength)),
    () => new Error("JSON fallback must not run"),
    () => binaryStatusResponse(403),
    () => binaryResponse(bytes, "text/plain", "browser-fetch"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.downloadFile(FILE_URL);

  assert.equal(result.method, "browser-fetch");
  assert.equal(transport.fallbackBinaryCalls, 1);
});

test("rejects oversized metadata before downloading bytes", async () => {
  const transport = new FakeBinaryTransport(
    () => jsonResponse(fileMetadata(MAX_DOWNLOAD_BYTES + 1)),
    () => new Error("fallback must not run"),
    () => new Error("binary must not run"),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  await assert.rejects(() => service.downloadFile(FILE_URL), /download limit/u);
  assert.equal(transport.primaryBinaryCalls, 0);
});

test("allows a larger source only for internal document processing", async () => {
  const bytes = new Uint8Array(MAX_DOWNLOAD_BYTES + 1);
  const transport = new FakeBinaryTransport(
    () => jsonResponse(fileMetadata(bytes.byteLength)),
    () => new Error("fallback must not run"),
    () => binaryResponse(bytes),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  const result = await service.downloadFile(
    FILE_URL,
    MAX_DOCUMENT_SOURCE_BYTES,
  );

  assert.equal(result.sizeBytes, bytes.byteLength);
  assert.deepEqual(
    transport.primaryBinaryLimits,
    [MAX_DOCUMENT_SOURCE_BYTES],
  );
});

test("rejects binary data whose size differs from metadata", async () => {
  const transport = new FakeBinaryTransport(
    () => jsonResponse(fileMetadata(5)),
    () => new Error("fallback must not run"),
    () => binaryResponse(new Uint8Array([1, 2])),
    () => new Error("binary fallback must not run"),
  );
  const service = new SharePointFileService(SITE_URL, transport);

  await assert.rejects(() => service.downloadFile(FILE_URL), /size did not match/u);
});

function resolve<T>(value: T | Error): T {
  if (value instanceof Error) {
    throw value;
  }
  return value;
}
