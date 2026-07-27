export const CURRENT_USER_API_PATH = "/_api/web/currentuser?$select=Id,Title,IsSiteAdmin";

export interface SharePointCurrentUser {
  readonly id: number;
  readonly displayName: string;
  readonly isSiteAdmin: boolean;
}

export function parseCurrentUserResponse(body: string): SharePointCurrentUser {
  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    throw new Error("SharePoint current-user response is not valid JSON.");
  }

  const root = asRecord(parsed, "SharePoint current-user response must be an object.");
  const payload = "d" in root ? asRecord(root.d, "SharePoint verbose response is invalid.") : root;

  const id = payload.Id;
  const displayName = payload.Title;
  const isSiteAdmin = payload.IsSiteAdmin;

  if (typeof id !== "number" || !Number.isSafeInteger(id) || id < 0) {
    throw new Error("SharePoint current-user response does not contain a valid Id.");
  }
  if (typeof displayName !== "string" || displayName.trim() === "") {
    throw new Error("SharePoint current-user response does not contain a valid Title.");
  }
  if (typeof isSiteAdmin !== "boolean") {
    throw new Error("SharePoint current-user response does not contain a valid IsSiteAdmin value.");
  }

  return {
    id,
    displayName: displayName.trim(),
    isSiteAdmin,
  };
}

function asRecord(value: unknown, message: string): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error(message);
  }
  return value as Record<string, unknown>;
}
