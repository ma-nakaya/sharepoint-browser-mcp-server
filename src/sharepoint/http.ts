export type RequestMethod = "api-request" | "browser-fetch";

export interface SharePointResponse {
  readonly status: number;
  readonly contentType: string;
  readonly body: string;
  readonly bodyTruncated?: boolean;
  readonly method: RequestMethod;
}

export interface SharePointTransport {
  get(apiPath: string): Promise<SharePointResponse>;
  getViaPage(apiPath: string): Promise<SharePointResponse>;
  close(): Promise<void>;
}

export interface SharePointBinaryResponse {
  readonly status: number;
  readonly contentType: string;
  readonly body: Uint8Array;
  readonly bodyTruncated?: boolean;
  readonly method: RequestMethod;
}

export interface SharePointBinaryTransport extends SharePointTransport {
  getBinary(
    apiPath: string,
    maxBytes?: number,
  ): Promise<SharePointBinaryResponse>;
  getBinaryViaPage(
    apiPath: string,
    maxBytes?: number,
  ): Promise<SharePointBinaryResponse>;
}
