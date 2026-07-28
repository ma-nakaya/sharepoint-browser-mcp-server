export function buildSharePointApiUrl(siteUrl: string, apiPath: string): URL {
  if (!apiPath.startsWith("/_api/")) {
    throw new Error("Only SharePoint REST paths under /_api/ are allowed.");
  }

  const site = new URL(siteUrl);
  const candidate = new URL(apiPath.slice(1), `${siteUrl}/`);
  assertAllowedSharePointApiUrl(candidate, site);
  return candidate;
}

export function assertAllowedSharePointApiUrl(candidate: URL, configuredSite: URL): void {
  if (candidate.protocol !== "https:" || candidate.origin !== configuredSite.origin) {
    throw new Error("SharePoint REST request must remain on the configured site origin.");
  }

  const sitePath = configuredSite.pathname.replace(/\/$/u, "");
  const expectedPrefix = `${sitePath}/_api/`;
  if (!candidate.pathname.startsWith(expectedPrefix)) {
    throw new Error("SharePoint REST request must remain under the configured site path.");
  }
}
