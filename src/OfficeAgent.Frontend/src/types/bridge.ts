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

export interface ChatMessage {
  id: string;
  role: string;
  content: string;
  createdAtUtc: string;
}

export interface ChatSession {
  id: string;
  title: string;
  createdAtUtc: string;
  updatedAtUtc: string;
  messages: ChatMessage[];
}

export interface SessionState {
  activeSessionId: string;
  sessions: ChatSession[];
}

export interface SelectionContext {
  hasSelection: boolean;
  workbookName: string;
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  isContiguous: boolean;
  headerPreview: string[];
  sampleRows: string[][];
  warningMessage?: string | null;
}

export interface ExcelCommand {
  commandType: string;
  sheetName?: string;
  targetAddress?: string;
  newSheetName?: string;
  values?: string[][];
  confirmed: boolean;
}

export interface ExcelCommandPreview {
  title: string;
  summary: string;
  details: string[];
}

export interface ExcelTableData {
  sheetName: string;
  address: string;
  headers: string[];
  rows: string[][];
}

export interface ExcelCommandResult {
  commandType: string;
  requiresConfirmation: boolean;
  status: string;
  message: string;
  preview?: ExcelCommandPreview;
  table?: ExcelTableData;
  selectionContext?: SelectionContext;
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

export interface BridgeEventEnvelope<TPayload = unknown> {
  type: string;
  payload?: TPayload;
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
