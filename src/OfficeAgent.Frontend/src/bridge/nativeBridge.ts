import type {
  AppSettings,
  BridgeErrorPayload,
  BridgeEventEnvelope,
  BridgeRequestEnvelope,
  BridgeResponseEnvelope,
  ExcelCommand,
  ExcelCommandResult,
  SelectionContext,
  SessionState,
  PingPayload,
  WebViewHostLike,
  WebViewMessageEventLike,
} from '../types/bridge';

const BRIDGE_TYPES = {
  ping: 'bridge.ping',
  getSettings: 'bridge.getSettings',
  getSelectionContext: 'bridge.getSelectionContext',
  selectionContextChanged: 'bridge.selectionContextChanged',
  getSessions: 'bridge.getSessions',
  saveSettings: 'bridge.saveSettings',
  executeExcelCommand: 'bridge.executeExcelCommand',
  runSkill: 'bridge.runSkill',
} as const;

const BROWSER_PREVIEW_PING: PingPayload = {
  host: 'browser-preview',
  version: 'dev',
};

const BROWSER_PREVIEW_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  model: 'gpt-5-mini',
};

const BROWSER_PREVIEW_SELECTION_CONTEXT: SelectionContext = {
  hasSelection: true,
  workbookName: 'Browser Preview.xlsx',
  sheetName: 'Sheet1',
  address: 'A1:C4',
  rowCount: 4,
  columnCount: 3,
  isContiguous: true,
  headerPreview: ['Name', 'Region', 'Amount'],
  sampleRows: [
    ['Project A', 'CN', '42'],
    ['Project B', 'US', '36'],
  ],
  warningMessage: null,
};

const BROWSER_PREVIEW_SESSIONS: SessionState = {
  activeSessionId: 'browser-preview-session',
  sessions: [
    {
      id: 'browser-preview-session',
      title: 'Browser preview',
      createdAtUtc: '2026-03-29T00:00:00.0000000Z',
      updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
      messages: [],
    },
  ],
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

type SelectionContextListener = (payload: SelectionContext) => void;

export class NativeBridge {
  private readonly webView?: WebViewHostLike;
  private readonly pendingRequests = new Map<string, PendingRequest>();
  private readonly selectionContextListeners = new Set<SelectionContextListener>();
  private readonly handleMessage = (event: WebViewMessageEventLike) => {
    const response = event.data;
    if (isBridgeEventEnvelope(response) && response.type === BRIDGE_TYPES.selectionContextChanged) {
      const payload = response.payload as SelectionContext | undefined;
      if (payload) {
        this.selectionContextListeners.forEach((listener) => listener(payload));
      }

      return;
    }

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

  getSettings() {
    return this.invoke<void, AppSettings>(BRIDGE_TYPES.getSettings);
  }

  getSelectionContext() {
    return this.invoke<void, SelectionContext>(BRIDGE_TYPES.getSelectionContext);
  }

  getSessions() {
    return this.invoke<void, SessionState>(BRIDGE_TYPES.getSessions);
  }

  saveSettings(payload: AppSettings) {
    return this.invoke<AppSettings, AppSettings>(BRIDGE_TYPES.saveSettings, payload);
  }

  executeExcelCommand(payload: ExcelCommand) {
    return this.invoke<ExcelCommand, ExcelCommandResult>(BRIDGE_TYPES.executeExcelCommand, payload);
  }

  runSkill(payload: unknown) {
    return this.invoke(BRIDGE_TYPES.runSkill, payload);
  }

  onSelectionContextChanged(listener: SelectionContextListener) {
    this.selectionContextListeners.add(listener);
    return () => {
      this.selectionContextListeners.delete(listener);
    };
  }

  private invoke<TPayload, TResult>(type: string, payload?: TPayload): Promise<TResult> {
    if (!this.webView) {
      if (type === BRIDGE_TYPES.ping) {
        return Promise.resolve(BROWSER_PREVIEW_PING as TResult);
      }

      if (type === BRIDGE_TYPES.getSettings) {
        return Promise.resolve(BROWSER_PREVIEW_SETTINGS as TResult);
      }

      if (type === BRIDGE_TYPES.getSelectionContext) {
        return Promise.resolve(BROWSER_PREVIEW_SELECTION_CONTEXT as TResult);
      }

      if (type === BRIDGE_TYPES.getSessions) {
        return Promise.resolve(BROWSER_PREVIEW_SESSIONS as TResult);
      }

      if (type === BRIDGE_TYPES.saveSettings) {
        return Promise.resolve({
          apiKey: typeof (payload as AppSettings | undefined)?.apiKey === 'string' ? (payload as AppSettings).apiKey : '',
          baseUrl: typeof (payload as AppSettings | undefined)?.baseUrl === 'string'
            ? (payload as AppSettings).baseUrl
            : BROWSER_PREVIEW_SETTINGS.baseUrl,
          model: typeof (payload as AppSettings | undefined)?.model === 'string'
            ? (payload as AppSettings).model
            : BROWSER_PREVIEW_SETTINGS.model,
        } as TResult);
      }

      if (type === BRIDGE_TYPES.executeExcelCommand) {
        try {
          return Promise.resolve(createBrowserPreviewCommandResult(validateBrowserPreviewCommand(payload as ExcelCommand)) as TResult);
        } catch (error) {
          return Promise.reject(error);
        }
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

function isBridgeEventEnvelope(value: unknown): value is BridgeEventEnvelope {
  if (!value || typeof value !== 'object') {
    return false;
  }

  const candidate = value as Record<string, unknown>;
  return typeof candidate.type === 'string' && !('requestId' in candidate) && !('ok' in candidate);
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

function createBrowserPreviewCommandResult(command: ExcelCommand): ExcelCommandResult {
  switch (command.commandType) {
    case 'excel.readSelectionTable':
      return {
        commandType: command.commandType,
        requiresConfirmation: false,
        status: 'completed',
        message: 'Read selection from Sheet1 A1:C4.',
        table: {
          sheetName: 'Sheet1',
          address: 'A1:C4',
          headers: ['Name', 'Region', 'Amount'],
          rows: [
            ['Project A', 'CN', '42'],
            ['Project B', 'US', '36'],
          ],
        },
        selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
      };
    case 'excel.addWorksheet':
      return createBrowserPreviewWriteResult(command, {
        previewTitle: 'Add worksheet',
        previewSummary: `Add worksheet "${command.newSheetName ?? 'New Sheet'}"`,
        completedMessage: `Worksheet "${command.newSheetName ?? 'New Sheet'}" created.`,
      });
    case 'excel.renameWorksheet':
      return createBrowserPreviewWriteResult(command, {
        previewTitle: 'Rename worksheet',
        previewSummary: `Rename worksheet "${command.sheetName ?? 'Sheet1'}" to "${command.newSheetName ?? 'Renamed Sheet'}"`,
        completedMessage: `Worksheet "${command.sheetName ?? 'Sheet1'}" renamed to "${command.newSheetName ?? 'Renamed Sheet'}".`,
      });
    case 'excel.deleteWorksheet':
      return createBrowserPreviewWriteResult(command, {
        previewTitle: 'Delete worksheet',
        previewSummary: `Delete worksheet "${command.sheetName ?? 'Sheet1'}"`,
        completedMessage: `Worksheet "${command.sheetName ?? 'Sheet1'}" deleted.`,
      });
    case 'excel.writeRange':
      return createBrowserPreviewWriteResult(command, {
        previewTitle: 'Write range',
        previewSummary: `Write ${(command.values ?? []).length} row(s) to ${command.targetAddress ?? 'A1'}`,
        completedMessage: `Wrote ${(command.values ?? []).length} row(s) to ${command.targetAddress ?? 'A1'}.`,
        details: (command.values ?? []).slice(0, 3).map((row) => row.join(' | ')),
      });
    default:
      throw new NativeBridgeError({
        code: 'bridge_unavailable',
        message: `Browser preview does not support ${command.commandType}.`,
      });
  }
}

function createBrowserPreviewWriteResult(
  command: ExcelCommand,
  options: {
    previewTitle: string;
    previewSummary: string;
    completedMessage: string;
    details?: string[];
  },
): ExcelCommandResult {
  if (!command.confirmed) {
    return {
      commandType: command.commandType,
      requiresConfirmation: true,
      status: 'preview',
      message: 'Confirm this Excel action before the workbook is modified.',
      preview: {
        title: options.previewTitle,
        summary: options.previewSummary,
        details: options.details ?? ['Workbook: Browser Preview.xlsx'],
      },
      selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
    };
  }

  return {
    commandType: command.commandType,
    requiresConfirmation: false,
    status: 'completed',
    message: options.completedMessage,
    selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
  };
}

function validateBrowserPreviewCommand(command: ExcelCommand): ExcelCommand {
  if (!command?.commandType) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Excel commands must include a commandType.',
    });
  }

  if (command.commandType === 'excel.writeRange') {
    if (!command.targetAddress?.trim()) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Write range commands require a target address.',
      });
    }

    const trimmedTargetAddress = command.targetAddress.trim();
    if (trimmedTargetAddress.includes('!')) {
      const [, targetRangeAddress = ''] = trimmedTargetAddress.split('!', 2);
      if (!targetRangeAddress.trim()) {
        throw new NativeBridgeError({
          code: 'invalid_command',
          message: 'Write range commands must include a cell reference in the target address.',
        });
      }
    }

    if (!command.values?.length || !command.values[0]?.length) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Write range commands require at least one row and one column of values.',
      });
    }

    const expectedColumnCount = command.values[0].length;
    if (command.values.some((row) => row.length !== expectedColumnCount)) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Write range commands require a rectangular values payload.',
      });
    }

    if (command.sheetName?.trim() && command.targetAddress.includes('!')) {
      const [qualifiedSheetName] = command.targetAddress.split('!', 1);
      if (
        qualifiedSheetName.trim() &&
        command.sheetName.trim().toLowerCase() !== qualifiedSheetName.trim().toLowerCase()
      ) {
        throw new NativeBridgeError({
          code: 'invalid_command',
          message: 'Write range commands cannot specify conflicting sheet names.',
        });
      }
    }
  }

  if (command.commandType === 'excel.addWorksheet' && !command.newSheetName?.trim()) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Add worksheet commands require a new sheet name.',
    });
  }

  if (command.commandType === 'excel.renameWorksheet') {
    if (!command.sheetName?.trim() || !command.newSheetName?.trim()) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Rename worksheet commands require both the current and new sheet names.',
      });
    }

    if (command.sheetName.trim().toLowerCase() === command.newSheetName.trim().toLowerCase()) {
      throw new NativeBridgeError({
        code: 'invalid_command',
        message: 'Rename worksheet commands must change the worksheet name.',
      });
    }
  }

  if (command.commandType === 'excel.deleteWorksheet' && !command.sheetName?.trim()) {
    throw new NativeBridgeError({
      code: 'invalid_command',
      message: 'Delete worksheet commands require a sheet name.',
    });
  }

  return command;
}
