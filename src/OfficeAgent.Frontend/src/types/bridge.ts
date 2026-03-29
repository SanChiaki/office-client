export interface BridgeRequestEnvelope<TPayload = unknown> {
  type: string;
  requestId: string;
  payload?: TPayload;
}

export interface BridgeErrorPayload {
  code: string;
  message: string;
}

export interface AppSettings {
  apiKey: string;
  baseUrl: string;
  model: string;
}

export interface BridgeResponseEnvelope<TPayload = unknown> {
  type: string;
  requestId: string;
  ok: boolean;
  payload?: TPayload;
  error?: BridgeErrorPayload;
}

export interface PingPayload {
  host: string;
  version: string;
}

export interface WebViewMessageEventLike {
  data: unknown;
}

export interface WebViewHostLike {
  addEventListener: (
    type: 'message',
    listener: (event: WebViewMessageEventLike) => void,
  ) => void;
  removeEventListener: (
    type: 'message',
    listener: (event: WebViewMessageEventLike) => void,
  ) => void;
  postMessage: (message: unknown) => void;
}

declare global {
  interface Window {
    chrome?: {
      webview?: WebViewHostLike;
    };
  }
}
