import type {
  BridgeErrorPayload,
  BridgeRequestEnvelope,
  BridgeResponseEnvelope,
  PingPayload,
  WebViewHostLike,
  WebViewMessageEventLike,
} from '../types/bridge';

const BRIDGE_TYPES = {
  ping: 'bridge.ping',
  getSelectionContext: 'bridge.getSelectionContext',
  getSessions: 'bridge.getSessions',
  saveSettings: 'bridge.saveSettings',
  executeExcelCommand: 'bridge.executeExcelCommand',
  runSkill: 'bridge.runSkill',
} as const;

const BROWSER_PREVIEW_PING: PingPayload = {
  host: 'browser-preview',
  version: 'dev',
};

class NativeBridgeError extends Error {
  public readonly code: string;

  constructor(error: BridgeErrorPayload) {
    super(error.message);
    this.code = error.code;
    this.name = 'NativeBridgeError';
  }
}

type PendingRequest = {
  resolve: (value: unknown) => void;
  reject: (reason?: unknown) => void;
};

export class NativeBridge {
  private readonly webView?: WebViewHostLike;
  private readonly pendingRequests = new Map<string, PendingRequest>();
  private readonly handleMessage = (event: WebViewMessageEventLike) => {
    const response = event.data;
    if (!isBridgeResponseEnvelope(response)) {
      return;
    }

    const pending = this.pendingRequests.get(response.requestId);
    if (!pending) {
      return;
    }

    this.pendingRequests.delete(response.requestId);

    if (response.ok) {
      pending.resolve(response.payload);
      return;
    }

    pending.reject(new NativeBridgeError(normalizeError(response.error)));
  };

  constructor(webView: WebViewHostLike | undefined = getWebViewHost()) {
    this.webView = webView;
    this.webView?.addEventListener('message', this.handleMessage);
  }

  dispose() {
    this.webView?.removeEventListener('message', this.handleMessage);
    this.pendingRequests.clear();
  }

  ping() {
    return this.invoke<void, PingPayload>(BRIDGE_TYPES.ping);
  }

  getSelectionContext() {
    return this.invoke(BRIDGE_TYPES.getSelectionContext, null);
  }

  getSessions() {
    return this.invoke(BRIDGE_TYPES.getSessions, null);
  }

  saveSettings(payload: unknown) {
    return this.invoke(BRIDGE_TYPES.saveSettings, payload);
  }

  executeExcelCommand(payload: unknown) {
    return this.invoke(BRIDGE_TYPES.executeExcelCommand, payload);
  }

  runSkill(payload: unknown) {
    return this.invoke(BRIDGE_TYPES.runSkill, payload);
  }

  private invoke<TPayload, TResult>(type: string, payload?: TPayload): Promise<TResult> {
    if (!this.webView) {
      if (type === BRIDGE_TYPES.ping) {
        return Promise.resolve(BROWSER_PREVIEW_PING as TResult);
      }

      return Promise.reject(
        new NativeBridgeError({
          code: 'bridge_unavailable',
          message: 'Native bridge is only available inside the Excel task pane.',
        }),
      );
    }

    const requestId = createRequestId();
    const request: BridgeRequestEnvelope<TPayload> = { type, requestId };
    if (payload !== undefined) {
      request.payload = payload;
    }

    return new Promise<TResult>((resolve, reject) => {
      this.pendingRequests.set(requestId, {
        resolve: (value) => resolve(value as TResult),
        reject,
      });
      this.webView?.postMessage(request);
    });
  }
}

export const nativeBridge = new NativeBridge();

function getWebViewHost() {
  return window.chrome?.webview;
}

function createRequestId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  return `req-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function isBridgeResponseEnvelope(value: unknown): value is BridgeResponseEnvelope {
  if (!value || typeof value !== 'object') {
    return false;
  }

  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.type === 'string' &&
    typeof candidate.requestId === 'string' &&
    typeof candidate.ok === 'boolean'
  );
}

function normalizeError(error: BridgeErrorPayload | undefined): BridgeErrorPayload {
  if (error?.code && error.message) {
    return error;
  }

  return {
    code: 'bridge_error',
    message: 'The native host returned an invalid error payload.',
  };
}
