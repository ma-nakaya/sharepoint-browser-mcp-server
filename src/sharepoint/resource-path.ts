export function encodeResourcePathArgument(serverRelativeUrl: string): string {
  const escapedPath = serverRelativeUrl.replaceAll("'", "''");
  return encodeURIComponent(escapedPath)
    .replaceAll("%2F", "/")
    .replaceAll("'", "%27");
}
