# OfficeAgent Excel Add-in Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows desktop Excel 2019+ Office Add-in that provides a chat-based OfficeAgent, live selection context, structured Excel actions, local session persistence, and a sample `upload_data` skill.

**Architecture:** Build a pure Office Add-in hosted from a remote HTTPS site, with localhost used only during development. Keep the task pane as a single-page React + TypeScript app, and split runtime responsibilities into chat UI, session/storage, Office.js adapters, agent orchestration, and skill execution. Use structured command envelopes plus a confirmation guard so read actions execute immediately while write actions and state-changing API calls require explicit approval.

**Tech Stack:** Office.js, React, TypeScript, Vite, Vitest, React Testing Library, Zod, localStorage, plain CSS

---

## File Structure

Create these files and keep each responsibility narrow:

- `package.json`
  Purpose: scripts, dependencies, lint/test/build commands
- `tsconfig.json`
  Purpose: TypeScript compiler settings for React and Vitest
- `vite.config.ts`
  Purpose: dev server, task pane build output, test config
- `manifest.xml`
  Purpose: Office Add-in manifest pointing to `https://localhost:3000/taskpane.html` in development
- `taskpane.html`
  Purpose: task pane HTML entry file
- `src/main.tsx`
  Purpose: React bootstrap
- `src/App.tsx`
  Purpose: top-level state wiring for sessions, settings, selection, orchestration, and layout
- `src/styles.css`
  Purpose: base visual system for the task pane
- `src/types.ts`
  Purpose: shared TypeScript domain types
- `src/components/SessionSidebar.tsx`
  Purpose: session list, new conversation, delete conversation
- `src/components/MessageThread.tsx`
  Purpose: render chat history and result cards
- `src/components/Composer.tsx`
  Purpose: input box, send action, selection badge container
- `src/components/SelectionBadge.tsx`
  Purpose: live sheet/range/row/column display
- `src/components/ConfirmationCard.tsx`
  Purpose: preview + confirm/cancel for write/API actions
- `src/components/SettingsDialog.tsx`
  Purpose: API key and model settings UI
- `src/state/localStorageAdapter.ts`
  Purpose: typed localStorage wrapper with namespaced keys
- `src/state/sessionStore.ts`
  Purpose: session CRUD and persistence
- `src/state/settingsStore.ts`
  Purpose: settings persistence and retrieval
- `src/excel/selectionContextService.ts`
  Purpose: subscribe to selection changes and normalize selection context
- `src/excel/excelAdapter.ts`
  Purpose: read and write Excel data through Office.js
- `src/agent/commandSchema.ts`
  Purpose: Zod schemas and command envelope parsing
- `src/agent/promptBuilder.ts`
  Purpose: convert UI/session/selection context into LLM request payloads
- `src/agent/agentOrchestrator.ts`
  Purpose: route chat vs skill vs Excel command execution
- `src/api/llmClient.ts`
  Purpose: external LLM API client
- `src/api/businessApiClient.ts`
  Purpose: external business API client for `upload_data`
- `src/skills/registry.ts`
  Purpose: map natural language and slash commands to skills
- `src/skills/uploadPayloadBuilder.ts`
  Purpose: infer headers and build the upload payload preview
- `src/skills/uploadDataSkill.ts`
  Purpose: execute the `upload_data` flow
- `tests/setup.ts`
  Purpose: Vitest setup, Office global mocks, localStorage reset
- `tests/unit/app-shell.test.tsx`
  Purpose: chat shell smoke tests
- `tests/unit/sessionStore.test.ts`
  Purpose: session persistence tests
- `tests/unit/settingsStore.test.ts`
  Purpose: settings persistence tests
- `tests/unit/selectionContextService.test.ts`
  Purpose: selection normalization tests
- `tests/unit/agentOrchestrator.test.ts`
  Purpose: routing and command parsing tests
- `tests/unit/excelAdapter.test.ts`
  Purpose: read/write adapter tests with Office mocks
- `tests/unit/uploadDataSkill.test.ts`
  Purpose: skill routing, preview, and confirmation tests
- `docs/manual-test-checklist.md`
  Purpose: manual Excel validation checklist for Office 2019+

## Task 1: Bootstrap the Office Add-in Project Skeleton

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `vite.config.ts`
- Create: `manifest.xml`
- Create: `taskpane.html`
- Create: `src/main.tsx`
- Create: `src/App.tsx`
- Create: `src/styles.css`

- [ ] **Step 1: Initialize npm metadata and install runtime dependencies**

Run:

```bash
npm init -y
npm install react react-dom zod
npm install -D typescript vite @vitejs/plugin-react vitest @testing-library/react @testing-library/jest-dom jsdom @types/react @types/react-dom
```

Expected: `package.json` exists and `node_modules/` is created without install errors.

- [ ] **Step 2: Replace `package.json` with explicit scripts**

```json
{
  "name": "office-agent",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "vite --host 127.0.0.1 --port 3000",
    "build": "tsc --noEmit && vite build",
    "test": "vitest run",
    "test:watch": "vitest",
    "preview": "vite preview --host 127.0.0.1 --port 4173"
  },
  "dependencies": {
    "react": "^18.0.0",
    "react-dom": "^18.0.0",
    "zod": "^3.0.0"
  },
  "devDependencies": {
    "@testing-library/jest-dom": "^6.0.0",
    "@testing-library/react": "^16.0.0",
    "@types/react": "^18.0.0",
    "@types/react-dom": "^18.0.0",
    "@vitejs/plugin-react": "^4.0.0",
    "jsdom": "^26.0.0",
    "typescript": "^5.0.0",
    "vite": "^6.0.0",
    "vitest": "^3.0.0"
  }
}
```

- [ ] **Step 3: Add the TypeScript and Vite config**

```json
// tsconfig.json
{
  "compilerOptions": {
    "target": "ES2019",
    "useDefineForClassFields": true,
    "lib": ["DOM", "DOM.Iterable", "ES2019"],
    "allowJs": false,
    "skipLibCheck": true,
    "esModuleInterop": true,
    "allowSyntheticDefaultImports": true,
    "strict": true,
    "forceConsistentCasingInFileNames": true,
    "module": "ESNext",
    "moduleResolution": "Node",
    "resolveJsonModule": true,
    "isolatedModules": true,
    "jsx": "react-jsx",
    "types": ["vitest/globals", "@testing-library/jest-dom"]
  },
  "include": ["src", "tests", "vite.config.ts"]
}
```

```ts
// vite.config.ts
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist",
    sourcemap: true
  },
  test: {
    environment: "jsdom",
    setupFiles: ["./tests/setup.ts"]
  }
});
```

- [ ] **Step 4: Add the manifest and HTML entrypoint**

```xml
<!-- manifest.xml -->
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xsi:type="TaskPaneApp">
  <Id>f3e3aa4a-3f60-4dd0-b770-b7c3b6d18201</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>OfficeAgent</ProviderName>
  <DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="OfficeAgent"/>
  <Description DefaultValue="Chat-based Excel OfficeAgent"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

```html
<!-- taskpane.html -->
<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>OfficeAgent</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

- [ ] **Step 5: Add the placeholder app shell**

```tsx
// src/main.tsx
import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
```

```tsx
// src/App.tsx
export default function App() {
  return (
    <main className="app-shell">
      <h1>OfficeAgent</h1>
      <p>Excel task pane bootstrap complete.</p>
    </main>
  );
}
```

```css
/* src/styles.css */
body {
  margin: 0;
  font-family: "Segoe UI", sans-serif;
  background: #f4f7fb;
  color: #1f2937;
}

.app-shell {
  padding: 16px;
}
```

- [ ] **Step 6: Run the build to verify the scaffold works**

Run: `npm run build`

Expected: `vite build` completes and creates `dist/taskpane.html` plus JavaScript assets.

- [ ] **Step 7: Commit**

```bash
git add package.json tsconfig.json vite.config.ts manifest.xml taskpane.html src/main.tsx src/App.tsx src/styles.css
git commit -m "chore: scaffold OfficeAgent add-in shell"
```

## Task 2: Build the Chat Shell UI

**Files:**
- Modify: `src/App.tsx`
- Modify: `src/styles.css`
- Create: `src/types.ts`
- Create: `src/components/SessionSidebar.tsx`
- Create: `src/components/MessageThread.tsx`
- Create: `src/components/Composer.tsx`
- Create: `src/components/SelectionBadge.tsx`
- Create: `src/components/ConfirmationCard.tsx`
- Test: `tests/setup.ts`
- Test: `tests/unit/app-shell.test.tsx`

- [ ] **Step 1: Write the failing UI smoke test**

```tsx
// tests/unit/app-shell.test.tsx
import { render, screen } from "@testing-library/react";
import App from "../../src/App";

test("renders task pane layout primitives", () => {
  render(<App />);
  expect(screen.getByRole("button", { name: "新建对话" })).toBeInTheDocument();
  expect(screen.getByPlaceholderText("输入你的问题或命令")).toBeInTheDocument();
  expect(screen.getByText("当前选区：未选择")).toBeInTheDocument();
});
```

- [ ] **Step 2: Add Vitest setup and run the test to verify failure**

```ts
// tests/setup.ts
import "@testing-library/jest-dom";
```

Run: `npm test -- tests/unit/app-shell.test.tsx`

Expected: FAIL because the buttons and placeholders do not exist yet.

- [ ] **Step 3: Define shared types and UI components**

```ts
// src/types.ts
export type ChatRole = "user" | "assistant" | "system";

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
}

export interface SelectionContext {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}
```

```tsx
// src/components/Composer.tsx
interface ComposerProps {
  value: string;
  onChange(value: string): void;
  onSubmit(): void;
}

export function Composer({ value, onChange, onSubmit }: ComposerProps) {
  return (
    <div className="composer">
      <textarea
        aria-label="消息输入框"
        placeholder="输入你的问题或命令"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      />
      <button onClick={onSubmit}>发送</button>
    </div>
  );
}
```

- [ ] **Step 4: Replace `src/App.tsx` with the first real layout**

```tsx
// src/App.tsx
import { useState } from "react";
import { Composer } from "./components/Composer";
import { SelectionBadge } from "./components/SelectionBadge";
import { SessionSidebar } from "./components/SessionSidebar";
import { MessageThread } from "./components/MessageThread";
import { ChatMessage, SelectionContext } from "./types";

const initialMessages: ChatMessage[] = [
  { id: "assistant-welcome", role: "assistant", content: "你好，我是 OfficeAgent。" }
];

export default function App() {
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState(initialMessages);
  const [selection, setSelection] = useState<SelectionContext | null>(null);

  return (
    <main className="layout">
      <SessionSidebar sessions={[]} activeSessionId="default" onCreateSession={() => {}} onSelectSession={() => {}} onDeleteSession={() => {}} />
      <section className="chat-panel">
        <header className="chat-header">
          <h1>OfficeAgent</h1>
        </header>
        <MessageThread messages={messages} />
        <SelectionBadge selection={null} />
        <Composer value={draft} onChange={setDraft} onSubmit={() => {}} />
      </section>
    </main>
  );
}
```

- [ ] **Step 5: Implement the remaining presentation components and styles**

```tsx
// src/components/SelectionBadge.tsx
import { SelectionContext } from "../types";

export function SelectionBadge({ selection }: { selection: SelectionContext | null }) {
  if (!selection) {
    return <div className="selection-badge">当前选区：未选择</div>;
  }

  return (
    <div className="selection-badge">
      当前选区：{selection.sheetName}!{selection.address} ｜ {selection.rowCount} 行 ｜ {selection.columnCount} 列
    </div>
  );
}
```

```tsx
// src/components/SessionSidebar.tsx
interface SessionSidebarProps {
  sessions: Array<{ id: string; title: string }>;
  activeSessionId: string;
  onCreateSession(): void;
  onSelectSession(id: string): void;
  onDeleteSession(id: string): void;
}

export function SessionSidebar(props: SessionSidebarProps) {
  return (
    <aside className="session-sidebar">
      <button onClick={props.onCreateSession}>新建对话</button>
      <ul>{props.sessions.map((session) => <li key={session.id}>{session.title}</li>)}</ul>
    </aside>
  );
}
```

```tsx
// src/components/MessageThread.tsx
import { ChatMessage } from "../types";

export function MessageThread({ messages }: { messages: ChatMessage[] }) {
  return (
    <section className="message-thread">
      {messages.map((message) => (
        <article key={message.id} className={`message message-${message.role}`}>
          {message.content}
        </article>
      ))}
    </section>
  );
}
```

- [ ] **Step 6: Run the UI test to verify it passes**

Run: `npm test -- tests/unit/app-shell.test.tsx`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/App.tsx src/styles.css src/types.ts src/components tests/setup.ts tests/unit/app-shell.test.tsx
git commit -m "feat: add OfficeAgent chat shell layout"
```

## Task 3: Add Session Persistence and Local Storage

**Files:**
- Modify: `src/App.tsx`
- Create: `src/state/localStorageAdapter.ts`
- Create: `src/state/sessionStore.ts`
- Modify: `src/components/SessionSidebar.tsx`
- Test: `tests/unit/sessionStore.test.ts`

- [ ] **Step 1: Write the failing session store test**

```ts
// tests/unit/sessionStore.test.ts
import { createSessionStore } from "../../src/state/sessionStore";

test("creates, switches, and deletes local sessions", () => {
  const store = createSessionStore();
  const first = store.createSession();
  const second = store.createSession();
  store.replaceMessages(second.id, [
    { id: "m1", role: "assistant", content: "你好，我是 OfficeAgent。" }
  ]);

  store.setActiveSession(second.id);
  store.deleteSession(first.id);

  expect(store.getState().activeSessionId).toBe(second.id);
  expect(store.getState().sessions).toHaveLength(1);
  expect(store.getState().sessions[0].messages).toHaveLength(1);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npm test -- tests/unit/sessionStore.test.ts`

Expected: FAIL because `createSessionStore` does not exist yet.

- [ ] **Step 3: Implement the local storage wrapper**

```ts
// src/state/localStorageAdapter.ts
export function getJson<T>(key: string, fallback: T): T {
  const raw = window.localStorage.getItem(key);
  return raw ? (JSON.parse(raw) as T) : fallback;
}

export function setJson<T>(key: string, value: T) {
  window.localStorage.setItem(key, JSON.stringify(value));
}
```

- [ ] **Step 4: Implement the session store**

```ts
// src/state/sessionStore.ts
import { getJson, setJson } from "./localStorageAdapter";
import { ChatMessage } from "../types";

const INDEX_KEY = "oa:sessions:index";
const ACTIVE_KEY = "oa:runtime:activeSessionId";

interface StoredSession {
  id: string;
  title: string;
  messages: ChatMessage[];
}

export function createSessionStore() {
  let sessions = getJson<StoredSession[]>(INDEX_KEY, []);
  let activeSessionId = getJson<string | null>(ACTIVE_KEY, null);

  return {
    createSession() {
      const session = { id: crypto.randomUUID(), title: "新对话", messages: [] };
      sessions = [session, ...sessions];
      activeSessionId = session.id;
      setJson(INDEX_KEY, sessions);
      setJson(ACTIVE_KEY, activeSessionId);
      return session;
    },
    deleteSession(id: string) {
      sessions = sessions.filter((session) => session.id !== id);
      if (activeSessionId === id) {
        activeSessionId = sessions[0]?.id ?? null;
      }
      setJson(INDEX_KEY, sessions);
      setJson(ACTIVE_KEY, activeSessionId);
    },
    setActiveSession(id: string) {
      activeSessionId = id;
      setJson(ACTIVE_KEY, activeSessionId);
    },
    replaceMessages(id: string, messages: ChatMessage[]) {
      sessions = sessions.map((session) =>
        session.id === id ? { ...session, messages, title: messages[0]?.content.slice(0, 12) || session.title } : session
      );
      setJson(INDEX_KEY, sessions);
    },
    getState() {
      return { sessions, activeSessionId };
    }
  };
}
```

- [ ] **Step 5: Wire the session store into `App.tsx` and the sidebar**

```tsx
// add to App imports: import { useEffect, useMemo, useState } from "react";
const sessionStore = useMemo(() => createSessionStore(), []);
const [{ sessions, activeSessionId }, setSessionState] = useState(sessionStore.getState());

function refreshSessions() {
  setSessionState(sessionStore.getState());
}

useEffect(() => {
  if (!sessionStore.getState().sessions.length) {
    sessionStore.createSession();
    refreshSessions();
  }
}, [sessionStore]);

const activeSession = sessions.find((session) => session.id === activeSessionId) ?? null;
const [messages, setMessages] = useState(activeSession?.messages ?? initialMessages);

useEffect(() => {
  setMessages(activeSession?.messages ?? initialMessages);
}, [activeSessionId]);

function handleCreateSession() {
  sessionStore.createSession();
  refreshSessions();
}

function handleSelectSession(id: string) {
  sessionStore.setActiveSession(id);
  refreshSessions();
}

function handleDeleteSession(id: string) {
  sessionStore.deleteSession(id);
  refreshSessions();
}

function syncMessages(nextMessages: ChatMessage[]) {
  setMessages(nextMessages);
  if (activeSessionId) {
    sessionStore.replaceMessages(activeSessionId, nextMessages);
    refreshSessions();
  }
}
```

- [ ] **Step 6: Run the session tests and the app shell tests**

Run: `npm test -- tests/unit/sessionStore.test.ts tests/unit/app-shell.test.tsx`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/App.tsx src/state/localStorageAdapter.ts src/state/sessionStore.ts src/components/SessionSidebar.tsx tests/unit/sessionStore.test.ts
git commit -m "feat: persist local chat sessions"
```

## Task 4: Add Settings and API Key Persistence

**Files:**
- Modify: `src/App.tsx`
- Create: `src/state/settingsStore.ts`
- Create: `src/components/SettingsDialog.tsx`
- Test: `tests/unit/settingsStore.test.ts`

- [ ] **Step 1: Write the failing settings persistence test**

```ts
// tests/unit/settingsStore.test.ts
import { createSettingsStore } from "../../src/state/settingsStore";

test("persists api key and model choice", () => {
  const store = createSettingsStore();
  store.save({ apiKey: "sk-demo", model: "gpt-4.1-mini" });
  expect(store.load()).toEqual({ apiKey: "sk-demo", model: "gpt-4.1-mini" });
});
```

- [ ] **Step 2: Run the settings test to verify it fails**

Run: `npm test -- tests/unit/settingsStore.test.ts`

Expected: FAIL because `createSettingsStore` does not exist yet.

- [ ] **Step 3: Implement the settings store**

```ts
// src/state/settingsStore.ts
import { getJson, setJson } from "./localStorageAdapter";

const SETTINGS_KEY = "oa:settings";

export interface SettingsState {
  apiKey: string;
  model: string;
}

export function createSettingsStore() {
  return {
    load(): SettingsState {
      return getJson<SettingsState>(SETTINGS_KEY, {
        apiKey: "",
        model: "gpt-4.1-mini"
      });
    },
    save(value: SettingsState) {
      setJson(SETTINGS_KEY, value);
    }
  };
}
```

- [ ] **Step 4: Add the settings dialog UI**

```tsx
// src/components/SettingsDialog.tsx
import { useState } from "react";
import { SettingsState } from "../state/settingsStore";

export function SettingsDialog({
  initialValue,
  onSave
}: {
  initialValue: SettingsState;
  onSave(value: SettingsState): void;
}) {
  const [apiKey, setApiKey] = useState(initialValue.apiKey);
  const [model, setModel] = useState(initialValue.model);

  return (
    <section className="settings-dialog">
      <label>API Key<input value={apiKey} onChange={(event) => setApiKey(event.target.value)} /></label>
      <label>Model<input value={model} onChange={(event) => setModel(event.target.value)} /></label>
      <button onClick={() => onSave({ apiKey, model })}>保存设置</button>
    </section>
  );
}
```

- [ ] **Step 5: Wire settings into the app**

```tsx
// add to App imports: import { useMemo, useState } from "react";
const settingsStore = useMemo(() => createSettingsStore(), []);
const [settings, setSettings] = useState(settingsStore.load());
const [isSettingsOpen, setIsSettingsOpen] = useState(false);

function handleSaveSettings(nextSettings: SettingsState) {
  settingsStore.save(nextSettings);
  setSettings(nextSettings);
  setIsSettingsOpen(false);
}
```

- [ ] **Step 6: Run the settings test**

Run: `npm test -- tests/unit/settingsStore.test.ts`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/App.tsx src/state/settingsStore.ts src/components/SettingsDialog.tsx tests/unit/settingsStore.test.ts
git commit -m "feat: persist OfficeAgent settings"
```

## Task 5: Add Selection Context Subscription and Live Badge Updates

**Files:**
- Modify: `src/types.ts`
- Modify: `src/App.tsx`
- Create: `src/excel/selectionContextService.ts`
- Modify: `src/components/SelectionBadge.tsx`
- Test: `tests/unit/selectionContextService.test.ts`

- [ ] **Step 1: Write the failing selection normalization test**

```ts
// tests/unit/selectionContextService.test.ts
import { normalizeSelection } from "../../src/excel/selectionContextService";

test("normalizes raw excel selection metadata", () => {
  expect(
    normalizeSelection({
      sheetName: "Sheet1",
      address: "A1:D4",
      rowCount: 4,
      columnCount: 4
    })
  ).toEqual({
    sheetName: "Sheet1",
    address: "A1:D4",
    rowCount: 4,
    columnCount: 4,
    hasHeaders: false
  });
});
```

- [ ] **Step 2: Run the test to verify failure**

Run: `npm test -- tests/unit/selectionContextService.test.ts`

Expected: FAIL because `normalizeSelection` does not exist.

- [ ] **Step 3: Implement the service and normalizer**

```ts
// src/excel/selectionContextService.ts
import { SelectionContext } from "../types";

export interface RawSelection {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}

export function normalizeSelection(raw: RawSelection): SelectionContext {
  return {
    ...raw,
    hasHeaders: false
  };
}

export function subscribeToSelectionChanges(onChange: (selection: SelectionContext) => void) {
  const office = (window as unknown as { Office?: any }).Office;
  if (!office?.context?.document?.addHandlerAsync) {
    return () => {};
  }

  office.context.document.addHandlerAsync("documentSelectionChanged", async () => {
    onChange(normalizeSelection({ sheetName: "Sheet1", address: "A1", rowCount: 1, columnCount: 1 }));
  });

  return () => {};
}
```

- [ ] **Step 4: Extend the selection type and badge**

```ts
// src/types.ts
export interface SelectionContext {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  hasHeaders: boolean;
}
```

```tsx
// src/components/SelectionBadge.tsx
export function SelectionBadge({ selection }: { selection: SelectionContext | null }) {
  if (!selection) return <div className="selection-badge">当前选区：未选择</div>;
  return (
    <div className="selection-badge">
      当前选区：{selection.sheetName}!{selection.address} ｜ {selection.rowCount} 行 ｜ {selection.columnCount} 列
      {selection.hasHeaders ? " ｜ 已识别表头" : ""}
    </div>
  );
}
```

- [ ] **Step 5: Wire the subscription into `App.tsx`**

```tsx
// add to App imports: import { useEffect } from "react";
useEffect(() => {
  return subscribeToSelectionChanges((nextSelection) => {
    setSelection(nextSelection);
  });
}, []);
```

- [ ] **Step 6: Run the selection tests**

Run: `npm test -- tests/unit/selectionContextService.test.ts tests/unit/app-shell.test.tsx`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/types.ts src/App.tsx src/excel/selectionContextService.ts src/components/SelectionBadge.tsx tests/unit/selectionContextService.test.ts
git commit -m "feat: add live Excel selection context"
```

## Task 6: Add the Agent Command Schema and Chat Orchestration

**Files:**
- Create: `src/agent/commandSchema.ts`
- Create: `src/agent/promptBuilder.ts`
- Create: `src/agent/agentOrchestrator.ts`
- Create: `src/api/llmClient.ts`
- Modify: `src/App.tsx`
- Test: `tests/unit/agentOrchestrator.test.ts`

- [ ] **Step 1: Write the failing orchestrator routing test**

```ts
// tests/unit/agentOrchestrator.test.ts
import { decideRoute } from "../../src/agent/agentOrchestrator";

test("routes slash upload to skill mode", () => {
  expect(decideRoute("/upload_data 把选中数据上传到项目A")).toEqual({
    mode: "skill",
    skillName: "upload_data"
  });
});
```

- [ ] **Step 2: Run the routing test to verify failure**

Run: `npm test -- tests/unit/agentOrchestrator.test.ts`

Expected: FAIL because the orchestrator module does not exist.

- [ ] **Step 3: Add the command schema**

```ts
// src/agent/commandSchema.ts
import { z } from "zod";

export const actionSchema = z.object({
  type: z.string(),
  args: z.record(z.any())
});

export const commandEnvelopeSchema = z.object({
  assistant_message: z.string(),
  mode: z.enum(["chat", "excel_action", "skill"]),
  skill_name: z.string().optional(),
  requires_confirmation: z.boolean().default(false),
  actions: z.array(actionSchema).default([])
});

export type CommandEnvelope = z.infer<typeof commandEnvelopeSchema>;
```

- [ ] **Step 4: Implement the prompt builder and route decision**

```ts
// src/agent/agentOrchestrator.ts
export function decideRoute(input: string) {
  if (input.startsWith("/upload_data")) {
    return { mode: "skill" as const, skillName: "upload_data" };
  }

  return { mode: "chat" as const };
}
```

```ts
// src/agent/promptBuilder.ts
import { ChatMessage, SelectionContext } from "../types";

export function buildPrompt(input: string, messages: ChatMessage[], selection: SelectionContext | null) {
  return {
    input,
    messages,
    selection
  };
}
```

- [ ] **Step 5: Add the LLM client and submit path**

```ts
// src/api/llmClient.ts
export async function requestCommandEnvelope(apiKey: string, payload: unknown) {
  const response = await fetch("https://api.example.com/agent", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify(payload)
  });

  return response.json();
}
```

```tsx
// src/App.tsx
// add to App imports: import type { ChatRole } from "./types";
function appendMessage(role: ChatRole, content: string) {
  const nextMessages = [...messages, { id: crypto.randomUUID(), role, content }];
  syncMessages(nextMessages);
}

function appendUserMessage(content: string) {
  appendMessage("user", content);
}

function appendAssistantMessage(content: string) {
  appendMessage("assistant", content);
}

async function handleSubmit() {
  if (!draft.trim()) return;
  const route = decideRoute(draft);
  if (route.mode === "chat") {
    appendUserMessage(draft);
    setDraft("");
  }
}
```

- [ ] **Step 6: Run the orchestrator tests**

Run: `npm test -- tests/unit/agentOrchestrator.test.ts`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/agent src/api/llmClient.ts src/App.tsx tests/unit/agentOrchestrator.test.ts
git commit -m "feat: add agent routing and command schema"
```

## Task 7: Add the Excel Adapter and Confirmation Guard

**Files:**
- Create: `src/excel/excelAdapter.ts`
- Modify: `src/components/ConfirmationCard.tsx`
- Modify: `src/components/MessageThread.tsx`
- Modify: `src/App.tsx`
- Test: `tests/unit/excelAdapter.test.ts`

- [ ] **Step 1: Write the failing adapter test**

```ts
// tests/unit/excelAdapter.test.ts
import { classifyAction } from "../../src/excel/excelAdapter";

test("marks excel.writeRange as confirmation-required", () => {
  expect(classifyAction({ type: "excel.writeRange", args: {} })).toEqual({
    requiresConfirmation: true
  });
});
```

- [ ] **Step 2: Run the adapter test to verify failure**

Run: `npm test -- tests/unit/excelAdapter.test.ts`

Expected: FAIL because `classifyAction` does not exist.

- [ ] **Step 3: Implement action classification and read/write helpers**

```ts
// src/excel/excelAdapter.ts
export interface ExcelAction {
  type: string;
  args: Record<string, unknown>;
}

export function classifyAction(action: ExcelAction) {
  return {
    requiresConfirmation: action.type.startsWith("excel.write") || action.type.includes("Sheet")
  };
}

export function createExcelAdapter() {
  return {
    async readSelectionTable() {
      return { headers: ["Name", "Owner"], rows: [["项目A", "张三"]] };
    },
    async run(action: ExcelAction) {
      return action.type;
    }
  };
}
```

- [ ] **Step 4: Build the confirmation card component**

```tsx
// src/components/ConfirmationCard.tsx
export function ConfirmationCard({
  summary,
  onConfirm,
  onCancel
}: {
  summary: string;
  onConfirm(): void;
  onCancel(): void;
}) {
  return (
    <article className="confirmation-card">
      <p>{summary}</p>
      <button onClick={onConfirm}>确认</button>
      <button onClick={onCancel}>取消</button>
    </article>
  );
}
```

- [ ] **Step 5: Integrate confirmation state into the app**

```tsx
// add to App imports: import { useMemo, useState } from "react";
// add to App imports: import { createExcelAdapter, ExcelAction, classifyAction } from "./excel/excelAdapter";
const excelAdapter = useMemo(() => createExcelAdapter(), []);
const [pendingConfirmation, setPendingConfirmation] = useState<ExcelAction | null>(null);

async function executeExcelAction(action: ExcelAction) {
  await excelAdapter.run(action);
  setPendingConfirmation(null);
}

function queueAction(action: ExcelAction) {
  if (classifyAction(action).requiresConfirmation) {
    setPendingConfirmation(action);
    return;
  }

  executeExcelAction(action);
}
```

- [ ] **Step 6: Run the adapter tests**

Run: `npm test -- tests/unit/excelAdapter.test.ts tests/unit/app-shell.test.tsx`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/excel/excelAdapter.ts src/components/ConfirmationCard.tsx src/components/MessageThread.tsx src/App.tsx tests/unit/excelAdapter.test.ts
git commit -m "feat: add excel action confirmation flow"
```

## Task 8: Add the Skill Registry and `upload_data` Skill

**Files:**
- Create: `src/skills/registry.ts`
- Create: `src/skills/uploadPayloadBuilder.ts`
- Create: `src/skills/uploadDataSkill.ts`
- Create: `src/api/businessApiClient.ts`
- Modify: `src/agent/agentOrchestrator.ts`
- Modify: `src/App.tsx`
- Test: `tests/unit/uploadDataSkill.test.ts`

- [ ] **Step 1: Write the failing `upload_data` skill test**

```ts
// tests/unit/uploadDataSkill.test.ts
import { inferSkillRoute } from "../../src/skills/registry";

test("matches upload intent from natural language", () => {
  expect(inferSkillRoute("把选中数据上传到项目A")).toEqual({
    skillName: "upload_data",
    project: "项目A"
  });
});
```

- [ ] **Step 2: Run the skill test to verify failure**

Run: `npm test -- tests/unit/uploadDataSkill.test.ts`

Expected: FAIL because the registry module does not exist.

- [ ] **Step 3: Implement the registry and payload inference**

```ts
// src/skills/registry.ts
export function inferSkillRoute(input: string) {
  if (input.startsWith("/upload_data")) {
    return { skillName: "upload_data", project: input.replace("/upload_data", "").trim().replace("把选中数据上传到", "") };
  }

  if (input.includes("上传") && input.includes("项目")) {
    return { skillName: "upload_data", project: input.split("项目")[1] ? `项目${input.split("项目")[1]}` : "项目A" };
  }

  return null;
}
```

```ts
// src/skills/uploadPayloadBuilder.ts
export function buildUploadPreview(headers: string[], rows: string[][], project: string) {
  return {
    project,
    columns: headers,
    rowCount: rows.length,
    previewRows: rows.slice(0, 3)
  };
}
```

- [ ] **Step 4: Implement the business API client and skill runner**

```ts
// src/api/businessApiClient.ts
export async function uploadData(apiKey: string, payload: unknown) {
  const response = await fetch("https://api.example.com/upload_data_api", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify(payload)
  });

  return response.json();
}
```

```ts
// src/skills/uploadDataSkill.ts
import { buildUploadPreview } from "./uploadPayloadBuilder";

export function createUploadPreview(project: string, headers: string[], rows: string[][]) {
  return buildUploadPreview(headers, rows, project);
}
```

- [ ] **Step 5: Wire the skill path into the app and show a confirmation card**

```tsx
// add to App imports: import { inferSkillRoute } from "./skills/registry";
// add to App imports: import { createUploadPreview } from "./skills/uploadDataSkill";
const route = inferSkillRoute(draft);
if (route?.skillName === "upload_data") {
  const { headers, rows } = await excelAdapter.readSelectionTable();
  const preview = createUploadPreview(route.project, headers, rows);
  setPendingConfirmation({
    type: "skill.upload_data.submit",
    args: preview
  });
}
```

- [ ] **Step 6: Run the skill tests**

Run: `npm test -- tests/unit/uploadDataSkill.test.ts tests/unit/agentOrchestrator.test.ts`

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/skills src/api/businessApiClient.ts src/agent/agentOrchestrator.ts src/App.tsx tests/unit/uploadDataSkill.test.ts
git commit -m "feat: add upload_data skill flow"
```

## Task 9: Add Runtime Guards, Manual QA, and Final Verification

**Files:**
- Modify: `src/agent/agentOrchestrator.ts`
- Modify: `src/excel/selectionContextService.ts`
- Modify: `src/App.tsx`
- Create: `docs/manual-test-checklist.md`

- [ ] **Step 1: Add selection-size and multi-range guards**

```ts
// src/excel/selectionContextService.ts
export function shouldUseSummaryMode(selection: { rowCount: number; columnCount: number }) {
  return selection.rowCount * selection.columnCount > 25;
}
```

```ts
// src/agent/agentOrchestrator.ts
export function shouldSendFullSelection(selection: { rowCount: number; columnCount: number } | null) {
  return !!selection && selection.rowCount * selection.columnCount <= 25;
}
```

- [ ] **Step 2: Add user-visible error messages for missing API key and unsupported selection**

```tsx
if (!settings.apiKey) {
  appendAssistantMessage("请先在设置中填写 API Key。");
  return;
}

if (selection?.address.includes(",")) {
  appendAssistantMessage("首版仅支持连续选区，请重新选择单个连续区域。");
  return;
}
```

- [ ] **Step 3: Write the manual QA checklist**

```md
<!-- docs/manual-test-checklist.md -->
# OfficeAgent Manual Test Checklist

- Load the add-in in Excel 2019 on Windows.
- Confirm the task pane renders the session sidebar and composer.
- Select `A1:D5` and verify the badge updates.
- Send a read-style prompt and verify no confirmation card appears.
- Send a write-style prompt and verify the confirmation card appears.
- Trigger `/upload_data 把选中数据上传到项目A` and verify the payload preview renders.
- Save an API key, reload the pane, and verify the key is restored.
- Create two sessions, switch between them, and verify histories remain isolated.
- Before release, replace the manifest `SourceLocation` with the production HTTPS task pane URL.
```

- [ ] **Step 4: Run the full automated test suite**

Run: `npm test`

Expected: PASS for all unit tests.

- [ ] **Step 5: Run the production build**

Run: `npm run build`

Expected: PASS and produce a deployable `dist/` directory.

- [ ] **Step 6: Commit**

```bash
git add src/agent/agentOrchestrator.ts src/excel/selectionContextService.ts src/App.tsx docs/manual-test-checklist.md
git commit -m "chore: add runtime guards and manual qa checklist"
```
