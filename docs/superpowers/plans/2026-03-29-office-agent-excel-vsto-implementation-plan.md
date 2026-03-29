# OfficeAgent Excel VSTO Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows Excel VSTO add-in that uses `WebView2 + React` for the task pane UI, supports local sessions and settings, reads live Excel context, executes Excel commands through native services, and ships an `upload_data` skill with confirmation.

**Architecture:** Use `VSTO` for the Excel host, Ribbon, task pane lifecycle, Excel Interop, local storage, and enterprise installation. Use `WebView2` to host a packaged React frontend inside the task pane, and define a typed JS/.NET bridge for all UI-to-native interactions. Keep business logic in Core services so the frontend remains a presentation layer and Excel COM access stays isolated in Infrastructure.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, WinForms, WebView2, React, TypeScript, Vite, HttpClient, JSON file storage, DPAPI, MSI packaging

---

## File Structure

- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Purpose: VSTO entrypoint, startup/shutdown, service bootstrap, Excel event registration
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Purpose: Ribbon tab/button to show or hide the task pane
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
  Purpose: WinForms `UserControl` that hosts WebView2
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
  Purpose: create/show/hide/synchronize the custom task pane
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
  Purpose: shared bridge request/response model
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
  Purpose: route web messages to native services
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
  Purpose: initialize WebView2, local content mapping, message wiring
- `src/OfficeAgent.Core/Models/*.cs`
  Purpose: session, message, selection, command, skill, preview models
- `src/OfficeAgent.Core/Services/IAgentOrchestrator.cs`
  Purpose: orchestration contract
- `src/OfficeAgent.Core/Services/IExcelContextService.cs`
  Purpose: selection and workbook context contract
- `src/OfficeAgent.Core/Services/IExcelCommandExecutor.cs`
  Purpose: Excel read/write command contract
- `src/OfficeAgent.Core/Services/ISessionStore.cs`
  Purpose: session persistence contract
- `src/OfficeAgent.Core/Services/ISettingsStore.cs`
  Purpose: settings persistence contract
- `src/OfficeAgent.Core/Services/ISkillRegistry.cs`
  Purpose: skill resolution contract
- `src/OfficeAgent.Core/Services/ConfirmationService.cs`
  Purpose: read/write confirmation policy
- `src/OfficeAgent.Core/Orchestration/AgentOrchestrator.cs`
  Purpose: route user input to chat, skill, or Excel command execution
- `src/OfficeAgent.Core/Skills/UploadDataSkill.cs`
  Purpose: `upload_data` multi-step workflow
- `src/OfficeAgent.Infrastructure/Excel/ExcelInteropAdapter.cs`
  Purpose: Excel Interop implementation
- `src/OfficeAgent.Infrastructure/Excel/ExcelSelectionContextService.cs`
  Purpose: translate Excel events into normalized selection context
- `src/OfficeAgent.Infrastructure/Http/LlmClient.cs`
  Purpose: LLM API access
- `src/OfficeAgent.Infrastructure/Http/BusinessApiClient.cs`
  Purpose: business API access
- `src/OfficeAgent.Infrastructure/Storage/FileSessionStore.cs`
  Purpose: JSON-backed session persistence
- `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`
  Purpose: JSON-backed settings persistence
- `src/OfficeAgent.Infrastructure/Security/DpapiSecretProtector.cs`
  Purpose: user-scope encryption for API key
- `src/OfficeAgent.Frontend/package.json`
  Purpose: frontend scripts and dependencies
- `src/OfficeAgent.Frontend/vite.config.ts`
  Purpose: frontend build config
- `src/OfficeAgent.Frontend/src/App.tsx`
  Purpose: task pane React shell
- `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
  Purpose: typed wrapper over `window.chrome.webview`
- `src/OfficeAgent.Frontend/src/components/*`
  Purpose: sidebar, thread, composer, settings, confirmation UI
- `src/OfficeAgent.Frontend/dist/*`
  Purpose: packaged task pane static assets
- `installer/OfficeAgent.Setup`
  Purpose: MSI packaging project
- `tests/OfficeAgent.Core.Tests/*`
  Purpose: core logic, routing, confirmation, payload, persistence tests
- `tests/OfficeAgent.Infrastructure.Tests/*`
  Purpose: file storage and HTTP tests
- `tests/OfficeAgent.Frontend.Tests/*`
  Purpose: frontend bridge and UI tests
- `docs/vsto-manual-test-checklist.md`
  Purpose: Excel desktop + installer verification checklist

## Task 1: Create the VSTO + Frontend Monorepo Skeleton

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn`
- Create: `src/OfficeAgent.Core`
- Create: `src/OfficeAgent.Infrastructure`
- Create: `src/OfficeAgent.Frontend`
- Create: `tests/OfficeAgent.Core.Tests`
- Create: `tests/OfficeAgent.Frontend.Tests`

- [ ] Create a new Excel VSTO Add-in solution targeting `.NET Framework 4.8`.
- [ ] Add class library projects for `OfficeAgent.Core` and `OfficeAgent.Infrastructure`, both targeting `net48`.
- [ ] Create the React frontend workspace under `src/OfficeAgent.Frontend`.
- [ ] Add a pure test project for core logic and storage code.
- [ ] Add frontend test tooling for bridge/UI code.
- [ ] Verify F5 can launch Excel with the empty add-in loaded.
- [ ] Commit with message: `chore: scaffold vsto and frontend solution`

## Task 2: Add Ribbon and Task Pane Host Shell

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Create: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- Create: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

- [ ] Add a Ribbon tab/group/button named `OfficeAgent`.
- [ ] Create `TaskPaneHostControl` as a WinForms `UserControl`.
- [ ] In `ThisAddIn_Startup`, initialize a singleton `TaskPaneController`.
- [ ] Make the Ribbon button show/hide the custom task pane.
- [ ] Set default right dock and stable width.
- [ ] Verify the task pane does not create duplicates after repeated toggles.
- [ ] Commit with message: `feat: add vsto ribbon and task pane shell`

## Task 3: Add WebView2 Host and Local Frontend Loading

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- Create: `src/OfficeAgent.Frontend/package.json`
- Create: `src/OfficeAgent.Frontend/vite.config.ts`
- Create: `src/OfficeAgent.Frontend/src/main.tsx`
- Create: `src/OfficeAgent.Frontend/src/App.tsx`

- [ ] Add the `Microsoft.Web.WebView2` SDK to the VSTO host project.
- [ ] Initialize a `WebView2` control inside `TaskPaneHostControl`.
- [ ] Build a minimal React frontend shell with:
  - session sidebar
  - message thread
  - composer
  - selection badge placeholder
  - settings button
- [ ] Package frontend assets locally and load them into WebView2.
- [ ] Use virtual host mapping instead of raw `file:///`.
- [ ] Verify task pane can render the React shell without any remote server dependency.
- [ ] Commit with message: `feat: host react task pane in webview2`

## Task 4: Define the JS/.NET Bridge Contract

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
- Create: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Create: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
- Create: `src/OfficeAgent.Frontend/src/types/bridge.ts`
- Test: `tests/OfficeAgent.Frontend.Tests/nativeBridge.test.ts`

- [ ] Define request/response envelopes with `type`, `requestId`, `payload`, `ok`, and `error`.
- [ ] Implement frontend wrapper methods:
  - `getSelectionContext`
  - `getSessions`
  - `saveSettings`
  - `executeExcelCommand`
  - `runSkill`
- [ ] Implement native message routing with a whitelist.
- [ ] Return structured errors instead of plain strings.
- [ ] Add tests for:
  - request/response correlation
  - unknown message rejection
  - malformed payload rejection
- [ ] Commit with message: `feat: add webview bridge contract`

## Task 5: Add Session and Settings Persistence

**Files:**
- Create: `src/OfficeAgent.Core/Models/ChatSession.cs`
- Create: `src/OfficeAgent.Core/Models/ChatMessage.cs`
- Create: `src/OfficeAgent.Core/Models/AppSettings.cs`
- Create: `src/OfficeAgent.Core/Services/ISessionStore.cs`
- Create: `src/OfficeAgent.Core/Services/ISettingsStore.cs`
- Create: `src/OfficeAgent.Infrastructure/Storage/FileSessionStore.cs`
- Create: `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`
- Create: `src/OfficeAgent.Infrastructure/Security/DpapiSecretProtector.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Test: `tests/OfficeAgent.Core.Tests/SessionStoreTests.cs`

- [ ] Store sessions under `%LocalAppData%\OfficeAgent\sessions\`.
- [ ] Store non-sensitive settings in `%LocalAppData%\OfficeAgent\settings.json`.
- [ ] Encrypt `API Key` with user-scope DPAPI before persistence.
- [ ] Include `Base URL` and `Model` in settings from the start.
- [ ] Add session operations:
  - create
  - switch
  - delete
  - restore last active
- [ ] Wire frontend settings and session sidebar to native persistence via the bridge.
- [ ] Add tests for malformed file recovery and secret roundtrip.
- [ ] Commit with message: `feat: persist sessions and settings through native stores`

## Task 6: Add Excel Selection Context Bridge

**Files:**
- Create: `src/OfficeAgent.Core/Models/SelectionContext.cs`
- Create: `src/OfficeAgent.Core/Services/IExcelContextService.cs`
- Create: `src/OfficeAgent.Infrastructure/Excel/ExcelSelectionContextService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Test: `tests/OfficeAgent.Core.Tests/SelectionContextTests.cs`

- [ ] Subscribe to `Application.SheetSelectionChange`.
- [ ] Normalize current selection into `SelectionContext`.
- [ ] Capture:
  - workbook name
  - sheet name
  - address
  - row/column count
  - contiguous/non-contiguous state
  - header preview
  - sample rows
- [ ] Push selection updates to the frontend through bridge events.
- [ ] Block unsupported multi-area selections with a clear user-facing message.
- [ ] Add tests for normalization logic independent of Excel runtime.
- [ ] Commit with message: `feat: bridge live excel selection context to web ui`

## Task 7: Add Excel Command Execution and Confirmation Flow

**Files:**
- Create: `src/OfficeAgent.Core/Models/ExcelCommand.cs`
- Create: `src/OfficeAgent.Core/Services/IExcelCommandExecutor.cs`
- Create: `src/OfficeAgent.Core/Services/ConfirmationService.cs`
- Create: `src/OfficeAgent.Infrastructure/Excel/ExcelInteropAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Test: `tests/OfficeAgent.Core.Tests/ConfirmationServiceTests.cs`

- [ ] Implement command classification:
  - read commands execute immediately
  - write commands require confirmation
- [ ] Add executor methods for:
  - read current selection as table
  - write range
  - add worksheet
  - rename worksheet
  - delete worksheet
- [ ] Return command previews to the frontend before write execution.
- [ ] Keep COM access entirely behind the executor service.
- [ ] Add tests for confirmation policy and command validation.
- [ ] Commit with message: `feat: add excel command execution and confirmation`

## Task 8: Add Agent Routing and upload_data Skill

**Files:**
- Create: `src/OfficeAgent.Core/Models/AgentCommandEnvelope.cs`
- Create: `src/OfficeAgent.Core/Services/IAgentOrchestrator.cs`
- Create: `src/OfficeAgent.Core/Services/ISkillRegistry.cs`
- Create: `src/OfficeAgent.Core/Orchestration/AgentOrchestrator.cs`
- Create: `src/OfficeAgent.Core/Skills/SkillRegistry.cs`
- Create: `src/OfficeAgent.Core/Skills/UploadDataSkill.cs`
- Create: `src/OfficeAgent.Core/Models/UploadPreview.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Test: `tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs`
- Test: `tests/OfficeAgent.Core.Tests/UploadDataSkillTests.cs`

- [ ] Preserve the current product rules:
  - natural language first
  - slash command forces the skill route
  - structured command envelope
  - read direct, write confirm
- [ ] Support routes:
  - chat
  - excel command
  - skill
- [ ] Add `upload_data` matching for both `/upload_data ...` and natural language.
- [ ] Read the current selection table through the native executor.
- [ ] Build preview payload from headers and sample rows.
- [ ] Return preview to the frontend, wait for confirmation, then call the business API.
- [ ] Commit with message: `feat: add upload_data skill over native bridge`

## Task 9: Add HTTP Clients and Configurable Base URL

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Http/LlmClient.cs`
- Create: `src/OfficeAgent.Infrastructure/Http/BusinessApiClient.cs`
- Modify: `src/OfficeAgent.Core/Models/AppSettings.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/BusinessApiClientTests.cs`

- [ ] Use `HttpClient` with configurable `Base URL`.
- [ ] Normalize `Base URL` by trimming whitespace and removing trailing slashes.
- [ ] Read protected `API Key` through the settings store.
- [ ] Add timeout, retry, and structured API error formatting.
- [ ] Add tests for:
  - default base URL fallback
  - trailing slash normalization
  - non-2xx error formatting
- [ ] Commit with message: `feat: add native http clients with configurable base url`

## Task 10: Package for Enterprise Distribution

**Files:**
- Create: `installer/OfficeAgent.Setup`
- Create: `docs/vsto-manual-test-checklist.md`
- Modify: solution packaging settings

- [ ] Build an MSI installer that deploys:
  - VSTO add-in
  - packaged frontend assets
  - WebView2 Runtime prerequisite handling
- [ ] Choose deployment mode:
  - default: bundle or invoke the Evergreen Standalone Installer for intranet/offline scenarios
- [ ] Add installer flows for:
  - install
  - uninstall
  - upgrade
- [ ] Verify add-in load behavior after fresh install and Excel restart.
- [ ] Verify no manifest side-load, catalog registration, or localhost certificate trust is needed for end users.
- [ ] Write a manual checklist covering:
  - Excel 2019 x86/x64
  - task pane open/close
  - selection updates
  - upload_data confirmation
  - API Key/Base URL save and reload
- [ ] Commit with message: `chore: package vsto add-in with webview2 runtime`

## Task 11: Stabilization and Release Readiness

**Files:**
- Modify: logging, diagnostics, settings recovery, docs as needed

- [ ] Add structured logging for:
  - startup
  - pane open/close
  - WebView2 initialization
  - bridge messages
  - selection changes
  - skill execution
  - HTTP failures
- [ ] Add defensive handling for:
  - protected worksheets
  - merged cells
  - invalid selection
  - workbook closed during pending action
  - WebView2 Runtime missing or initialization failure
- [ ] Run the full test suite.
- [ ] Run manual QA in supported Excel environments.
- [ ] Produce a release note and known-limitations document.
- [ ] Commit with message: `chore: harden vsto webview officeagent mvp`

