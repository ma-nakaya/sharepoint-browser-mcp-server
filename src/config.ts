import os from "node:os";
import path from "node:path";

export const BROWSER_CHANNELS = [
  "msedge",
  "msedge-beta",
  "msedge-dev",
  "msedge-canary",
] as const;

export type BrowserChannel = (typeof BROWSER_CHANNELS)[number];

export interface AppConfig {
  readonly siteUrl: string;
  readonly siteOrigin: string;
  readonly profileDir: string;
  readonly browserChannel: BrowserChannel;
  readonly headless: boolean;
  readonly requestTimeoutMs: number;
}

export function loadConfig(
  env: NodeJS.ProcessEnv = process.env,
  options: { readonly forceHeaded?: boolean } = {},
): AppConfig {
  const rawSiteUrl = requireValue(env.SHAREPOINT_SITE_URL, "SHAREPOINT_SITE_URL");
  const siteUrl = normalizeSharePointSiteUrl(rawSiteUrl);
  const profileDir = resolveProfileDir(env.SHAREPOINT_PROFILE_DIR, env);
  assertSafeProfileDirectory(profileDir);

  const browserChannel = parseBrowserChannel(env.SHAREPOINT_BROWSER_CHANNEL);
  const configuredHeadless = parseBoolean(env.SHAREPOINT_HEADLESS, true, "SHAREPOINT_HEADLESS");
  const requestTimeoutMs = parsePositiveInteger(
    env.SHAREPOINT_REQUEST_TIMEOUT_MS,
    15_000,
    "SHAREPOINT_REQUEST_TIMEOUT_MS",
  );

  return {
    siteUrl,
    siteOrigin: new URL(siteUrl).origin,
    profileDir,
    browserChannel,
    headless: options.forceHeaded === true ? false : configuredHeadless,
    requestTimeoutMs,
  };
}

export function normalizeSharePointSiteUrl(value: string): string {
  let url: URL;
  try {
    url = new URL(value.trim());
  } catch {
    throw new Error("SHAREPOINT_SITE_URL must be an absolute URL.");
  }

  if (url.protocol !== "https:") {
    throw new Error("SHAREPOINT_SITE_URL must use HTTPS.");
  }
  if (url.username !== "" || url.password !== "") {
    throw new Error("SHAREPOINT_SITE_URL must not contain credentials.");
  }
  if (url.search !== "" || url.hash !== "") {
    throw new Error("SHAREPOINT_SITE_URL must not contain a query string or fragment.");
  }

  const hostname = url.hostname.toLowerCase();
  if (!hostname.endsWith(".sharepoint.com")) {
    throw new Error("SHAREPOINT_SITE_URL must use a SharePoint Online host.");
  }
  if (hostname.endsWith("-my.sharepoint.com")) {
    throw new Error("OneDrive hosts are not allowed.");
  }

  const segments = url.pathname.split("/").filter(Boolean);
  const managedPath = segments[0]?.toLowerCase();
  if ((managedPath !== "sites" && managedPath !== "teams") || segments.length < 2) {
    throw new Error(
      "SHAREPOINT_SITE_URL must identify a site collection under /sites/<name> or /teams/<name>.",
    );
  }

  url.hostname = hostname;
  url.pathname = `/${segments.join("/")}`;
  return url.toString().replace(/\/$/, "");
}

export function resolveProfileDir(
  configuredValue: string | undefined,
  env: NodeJS.ProcessEnv = process.env,
): string {
  const configured = configuredValue?.trim();
  if (configured) {
    return path.resolve(configured);
  }

  const baseDir = env.LOCALAPPDATA?.trim() || path.join(os.homedir(), ".local", "share");
  return path.resolve(baseDir, "sharepoint-browser-mcp-server", "edge-profile");
}

export function assertSafeProfileDirectory(profileDir: string): void {
  const resolved = path.resolve(profileDir);
  const normalized = resolved.toLowerCase();
  const basename = path.basename(resolved).toLowerCase();
  const protectedPaths = new Set([
    path.parse(resolved).root.toLowerCase(),
    path.resolve(os.homedir()).toLowerCase(),
    path.resolve(process.cwd()).toLowerCase(),
  ]);

  if (protectedPaths.has(normalized)) {
    throw new Error("SHAREPOINT_PROFILE_DIR must be a dedicated subdirectory.");
  }

  if (basename === "user data" || basename === "default" || /^profile\s+\d+$/u.test(basename)) {
    throw new Error("SHAREPOINT_PROFILE_DIR must not point to a normal Edge profile.");
  }

  const normalizedWithSlashes = normalized.replaceAll("\\", "/");
  const edgeUserDataPattern = /\/microsoft\/edge(?: beta| dev| sxs)?\/user data(?:\/|$)/u;
  if (edgeUserDataPattern.test(normalizedWithSlashes)) {
    throw new Error("SHAREPOINT_PROFILE_DIR must not be inside Edge's normal User Data directory.");
  }
}

function requireValue(value: string | undefined, name: string): string {
  const normalized = value?.trim();
  if (!normalized) {
    throw new Error(`${name} is required.`);
  }
  return normalized;
}

function parseBrowserChannel(value: string | undefined): BrowserChannel {
  const normalized = value?.trim() || "msedge";
  if (!BROWSER_CHANNELS.includes(normalized as BrowserChannel)) {
    throw new Error(
      `SHAREPOINT_BROWSER_CHANNEL must be one of: ${BROWSER_CHANNELS.join(", ")}.`,
    );
  }
  return normalized as BrowserChannel;
}

function parseBoolean(value: string | undefined, defaultValue: boolean, name: string): boolean {
  const normalized = value?.trim().toLowerCase();
  if (!normalized) {
    return defaultValue;
  }
  if (normalized === "true") {
    return true;
  }
  if (normalized === "false") {
    return false;
  }
  throw new Error(`${name} must be true or false.`);
}

function parsePositiveInteger(value: string | undefined, defaultValue: number, name: string): number {
  const normalized = value?.trim();
  if (!normalized) {
    return defaultValue;
  }

  const parsed = Number(normalized);
  if (!Number.isSafeInteger(parsed) || parsed <= 0) {
    throw new Error(`${name} must be a positive integer.`);
  }
  return parsed;
}
