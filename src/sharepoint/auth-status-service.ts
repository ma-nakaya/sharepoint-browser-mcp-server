import {
  CURRENT_USER_API_PATH,
  parseCurrentUserResponse,
  type SharePointCurrentUser,
} from "./current-user.js";
import type {
  RequestMethod,
  SharePointResponse,
  SharePointTransport,
} from "./http.js";

export const AUTH_STATES = [
  "AUTHENTICATED",
  "LOGIN_REQUIRED",
  "ACCESS_DENIED",
  "SITE_NOT_FOUND",
  "UNAVAILABLE",
] as const;

export type AuthState = (typeof AUTH_STATES)[number];

export interface AuthStatus {
  readonly authenticated: boolean;
  readonly state: AuthState;
  readonly siteUrl: string;
  readonly message: string;
  readonly method?: RequestMethod;
  readonly user?: SharePointCurrentUser;
}

type Classification =
  | { readonly kind: "final"; readonly result: AuthStatus }
  | { readonly kind: "fallback" };

export class AuthStatusService {
  constructor(
    private readonly siteUrl: string,
    private readonly transport: SharePointTransport,
  ) {}

  async getStatus(): Promise<AuthStatus> {
    try {
      const primaryResponse = await this.transport.get(CURRENT_USER_API_PATH);
      const primary = this.classify(primaryResponse, true);
      if (primary.kind === "final") {
        return primary.result;
      }
    } catch {
      // Browser-page fetch is the fallback for session/proxy differences.
    }

    try {
      const browserResponse = await this.transport.getViaPage(CURRENT_USER_API_PATH);
      const browser = this.classify(browserResponse, false);
      if (browser.kind === "final") {
        return browser.result;
      }
      return this.unavailable("SharePoint returned an unexpected authentication response.");
    } catch {
      return this.unavailable("SharePoint could not be reached through the authenticated Edge session.");
    }
  }

  private classify(response: SharePointResponse, allowFallback: boolean): Classification {
    if (response.status === 200) {
      if (!isJsonContentType(response.contentType)) {
        return allowFallback
          ? { kind: "fallback" }
          : { kind: "final", result: this.unavailable("SharePoint returned a non-JSON response.") };
      }

      try {
        const user = parseCurrentUserResponse(response.body);
        return {
          kind: "final",
          result: {
            authenticated: true,
            state: "AUTHENTICATED",
            siteUrl: this.siteUrl,
            method: response.method,
            user,
            message: "The Edge session is authenticated for the configured SharePoint site.",
          },
        };
      } catch {
        return allowFallback
          ? { kind: "fallback" }
          : {
              kind: "final",
              result: this.unavailable("SharePoint returned an unexpected current-user response."),
            };
      }
    }

    if (response.status === 401 || isRedirect(response.status) || response.status === 0) {
      return allowFallback
        ? { kind: "fallback" }
        : {
            kind: "final",
            result: {
              authenticated: false,
              state: "LOGIN_REQUIRED",
              siteUrl: this.siteUrl,
              method: response.method,
              message: "Open the dedicated Edge profile and sign in to SharePoint.",
            },
          };
    }

    if (response.status === 403) {
      if (allowFallback) {
        return { kind: "fallback" };
      }
      return {
        kind: "final",
        result: {
          authenticated: false,
          state: "ACCESS_DENIED",
          siteUrl: this.siteUrl,
          method: response.method,
          message: "The signed-in user does not have access to the configured SharePoint site.",
        },
      };
    }

    if (response.status === 404) {
      return {
        kind: "final",
        result: {
          authenticated: false,
          state: "SITE_NOT_FOUND",
          siteUrl: this.siteUrl,
          method: response.method,
          message: "The configured SharePoint site was not found.",
        },
      };
    }

    return {
      kind: "final",
      result: this.unavailable(`SharePoint returned HTTP ${response.status}.`, response.method),
    };
  }

  private unavailable(message: string, method?: RequestMethod): AuthStatus {
    return {
      authenticated: false,
      state: "UNAVAILABLE",
      siteUrl: this.siteUrl,
      ...(method ? { method } : {}),
      message,
    };
  }
}

function isJsonContentType(contentType: string): boolean {
  const normalized = contentType.toLowerCase();
  return normalized.includes("application/json") || normalized.includes("text/json");
}

function isRedirect(status: number): boolean {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}
