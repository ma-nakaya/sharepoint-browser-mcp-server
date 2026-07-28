import type { RequestMethod } from "./http.js";
import { canvasHtmlToPlainText } from "./page-content.js";

export const MAX_LIST_ITEM_TEXT_CHARACTERS = 50_000;

export interface NormalizedSharePointListItem {
  readonly itemUrl: string;
  readonly serverRelativeListUrl: string;
  readonly itemId: number;
}

export interface SharePointListItemContent {
  readonly title: string;
  readonly url: string;
  readonly serverRelativeListUrl: string;
  readonly itemId: number;
  readonly modifiedTime?: string;
  readonly text: string;
  readonly fieldCount: number;
  readonly truncated: boolean;
  readonly method: RequestMethod;
}

interface ListItemTextField {
  readonly label: string;
  readonly value: string;
}

type JsonRecord = Record<string, unknown>;

const SYSTEM_FIELD_NAMES = new Set([
  "appauthor",
  "appeditor",
  "attachments",
  "author",
  "contenttype",
  "contenttypeid",
  "created",
  "created_x0020_date",
  "docicon",
  "edit",
  "editor",
  "filedirref",
  "fileleafref",
  "fileref",
  "folderchildcount",
  "fsobjtype",
  "guid",
  "id",
  "instanceid",
  "itemchildcount",
  "last_x0020_modified",
  "modified",
  "order",
  "owshiddenversion",
  "permask",
  "selecttitle",
  "title",
  "uniqueid",
  "workflowversion",
  "_level",
  "_moderationstatus",
  "_uiversionstring",
]);

const NORMALIZED_SYSTEM_FIELD_NAMES = new Set([
  "accesspolicy",
  "appauthor",
  "appeditor",
  "attachments",
  "author",
  "colorhex",
  "colortag",
  "commentcount",
  "commentflags",
  "complianceassetid",
  "complianceflags",
  "compliancetag",
  "compliancetaguserid",
  "compliancetagwrittentime",
  "contentversion",
  "contenttypeid",
  "copysource",
  "created",
  "createddate",
  "draftownerid",
  "edit",
  "editor",
  "emoji",
  "filedirref",
  "fileleafref",
  "fileref",
  "filetype",
  "folderchildcount",
  "fsobjtype",
  "guid",
  "hascopydestinations",
  "id",
  "instanceid",
  "iscurrentversion",
  "itemchildcount",
  "lastmodified",
  "level",
  "metainfo",
  "modified",
  "moderationcomments",
  "moderationstatus",
  "noexecute",
  "order",
  "originatorid",
  "owshiddenversion",
  "parentuniqueid",
  "permask",
  "progid",
  "ransomwareanomalymetainfo",
  "restricted",
  "selecttitle",
  "scopeid",
  "smlastmodifieddate",
  "smtotalfilecount",
  "smtotalfilestreamsize",
  "smtotalsize",
  "sortbehavior",
  "syncclientid",
  "taxcatchall",
  "title",
  "uniqueid",
  "uiversion",
  "uiversionstring",
  "virusinfo",
  "virusstatus",
  "virusvendorid",
  "workflowinstanceid",
  "workflowversion",
]);

export function isSharePointListItemFormUrl(value: string): boolean {
  const normalized = value.trim();
  if (!normalized) {
    return false;
  }
  try {
    const candidate = normalized.startsWith("/")
      ? new URL(normalized, "https://sharepoint.invalid")
      : new URL(normalized);
    return /\/Lists\/[^/]+\/DispForm\.aspx$/iu.test(candidate.pathname);
  } catch {
    return false;
  }
}

export function normalizeSharePointListItemUrl(
  siteUrl: string,
  itemUrl: string,
): NormalizedSharePointListItem {
  const normalizedInput = itemUrl.trim();
  if (!normalizedInput) {
    throw new Error("List item URL must not be empty.");
  }

  const site = new URL(siteUrl);
  let candidate: URL;
  try {
    candidate = normalizedInput.startsWith("/")
      ? new URL(normalizedInput, site.origin)
      : new URL(normalizedInput);
  } catch {
    throw new Error("List item URL must be an absolute or server-relative URL.");
  }

  if (
    candidate.protocol !== "https:" ||
    candidate.origin !== site.origin ||
    candidate.username !== "" ||
    candidate.password !== ""
  ) {
    throw new Error(
      "List item URL must remain on the configured SharePoint origin.",
    );
  }

  const decodedPath = decodeUrlPath(candidate.pathname, "List item");
  const sitePath = decodeUrlPath(site.pathname, "Configured site").replace(
    /\/$/u,
    "",
  );
  const listPrefix = `${sitePath}/Lists/`;
  if (!decodedPath.toLowerCase().startsWith(listPrefix.toLowerCase())) {
    throw new Error(
      "List item URL must identify a display form under the configured site's Lists path.",
    );
  }

  const listRelativePath = decodedPath.slice(listPrefix.length);
  const segments = listRelativePath.split("/");
  if (
    segments.length !== 2 ||
    !segments[0] ||
    segments[0] === "." ||
    segments[0] === ".." ||
    segments[1]?.toLowerCase() !== "dispform.aspx"
  ) {
    throw new Error(
      "List item URL must identify a DispForm.aspx page directly under a SharePoint list.",
    );
  }

  const idValues = [...candidate.searchParams.entries()]
    .filter(([key]) => key.toLowerCase() === "id")
    .map(([, value]) => value.trim());
  if (idValues.length !== 1 || !/^[1-9][0-9]*$/u.test(idValues[0] ?? "")) {
    throw new Error(
      "List item URL must contain exactly one positive integer ID parameter.",
    );
  }
  const itemId = Number(idValues[0]);
  if (!Number.isSafeInteger(itemId)) {
    throw new Error("List item ID exceeds the supported integer range.");
  }

  const serverRelativeListUrl = decodedPath.slice(
    0,
    -"/DispForm.aspx".length,
  );
  const canonical = new URL(site.origin);
  canonical.pathname = decodedPath;
  canonical.searchParams.set("ID", String(itemId));

  return {
    itemUrl: canonical.toString(),
    serverRelativeListUrl,
    itemId,
  };
}

export function buildListItemContentApiPath(
  item: NormalizedSharePointListItem,
): string {
  const parameters = new URLSearchParams({
    "@listUrl": `'${escapeODataString(item.serverRelativeListUrl)}'`,
  });
  return [
    `/_api/web/GetList(@listUrl)/items(${item.itemId})/FieldValuesAsText`,
    `?${parameters.toString()}`,
  ].join("");
}

export function parseListItemContentResponse(
  body: string,
  item: NormalizedSharePointListItem,
  method: RequestMethod,
): SharePointListItemContent {
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("SharePoint returned malformed list-item JSON.");
  }

  const root = asRecord(parsed);
  const envelope = asRecord(root?.d) ?? root;
  const fields = asRecord(envelope?.FieldValuesAsText) ?? envelope;
  if (!fields) {
    throw new Error("SharePoint returned an unexpected list-item response.");
  }

  const title =
    cleanFieldValue(fields.Title) ??
    cleanFieldValue(fields.LinkTitle) ??
    cleanFieldValue(fields.FileLeafRef) ??
    `List item ${item.itemId}`;
  const modifiedTime =
    cleanMetadataValue(fields.Modified) ??
    cleanMetadataValue(fields.Last_x0020_Modified);
  const textFields = collectTextFields(fields, title);
  const fullText = [
    title,
    ...textFields.map(({ label, value }) =>
      value.includes("\n") ? `${label}:\n${value}` : `${label}: ${value}`,
    ),
  ].join("\n\n");
  const truncated = fullText.length > MAX_LIST_ITEM_TEXT_CHARACTERS;

  return {
    title,
    url: item.itemUrl,
    serverRelativeListUrl: item.serverRelativeListUrl,
    itemId: item.itemId,
    ...(modifiedTime ? { modifiedTime } : {}),
    text: truncated
      ? fullText.slice(0, MAX_LIST_ITEM_TEXT_CHARACTERS)
      : fullText,
    fieldCount: textFields.length,
    truncated,
    method,
  };
}

function collectTextFields(
  fields: JsonRecord,
  title: string,
): ListItemTextField[] {
  const result: ListItemTextField[] = [];
  const seenValues = new Set([title]);

  for (const [internalName, rawValue] of Object.entries(fields)) {
    const normalizedName = internalName.toLowerCase();
    const label = decodeSharePointFieldName(internalName);
    const normalizedLabel = label
      .replace(/[^a-z0-9]/giu, "")
      .toLowerCase();
    if (
      internalName.startsWith("@") ||
      internalName.startsWith("__") ||
      SYSTEM_FIELD_NAMES.has(normalizedName) ||
      NORMALIZED_SYSTEM_FIELD_NAMES.has(normalizedLabel) ||
      /^[0-9a-f]{32}$/iu.test(label)
    ) {
      continue;
    }

    const value = cleanFieldValue(rawValue);
    if (!value || seenValues.has(value)) {
      continue;
    }
    seenValues.add(value);
    result.push({
      label,
      value,
    });
  }

  return result;
}

function cleanFieldValue(value: unknown): string | undefined {
  const scalar =
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
      ? String(value)
      : undefined;
  if (!scalar) {
    return undefined;
  }

  const normalized = canvasHtmlToPlainText(
    scalar.replace(/;#/gu, "; ").replace(/\p{Cf}/gu, ""),
  ).trim();
  return normalized || undefined;
}

function cleanMetadataValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim()
    ? value.trim().slice(0, 100)
    : undefined;
}

function decodeSharePointFieldName(value: string): string {
  let decoded = value;
  for (let pass = 0; pass < 3; pass += 1) {
    const next = decoded.replace(
      /_?x([0-9a-f]{4})_/giu,
      (_match, code: string) =>
        String.fromCharCode(Number.parseInt(code, 16)),
    );
    if (next === decoded) {
      break;
    }
    decoded = next;
  }

  return decoded
    .replace(/^OData_+/iu, "")
    .replace(/_+/gu, " ")
    .replace(/(\p{Script=Han})\s+(?=\p{Script=Han})/gu, "$1")
    .trim();
}

function decodeUrlPath(value: string, fieldName: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    throw new Error(`${fieldName} URL contains invalid percent encoding.`);
  }
}

function escapeODataString(value: string): string {
  return value.replaceAll("'", "''");
}

function asRecord(value: unknown): JsonRecord | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : undefined;
}
