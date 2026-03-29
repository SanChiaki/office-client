# OfficeAgent Excel VSTO Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows Excel VSTO add-in that provides a chat-based OfficeAgent task pane, live Excel context, local session persistence, external API access, and an `upload_data` skill with confirmation.

**Architecture:** Use a `VSTO Excel Add-in` as the host, expose a Ribbon button to open a `CustomTaskPane`, render the chat UI with WPF, and isolate business logic from Excel interop through service interfaces. Store sessions locally on disk, protect secrets with DPAPI, and package the add-in as an MSI for enterprise distribution.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, WPF, WinForms `CustomTaskPane`, `HttpClient`, JSON file storage, DPAPI, MSTest or xUnit, MSI packaging

---

## File Structure

- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Purpose: VSTO entrypoint, startup/shutdown, service bootstrap, Excel event registration
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Purpose: Ribbon tab/button to show or hide the task pane
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
  Purpose: WinForms host for the custom task pane
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
  Purpose: create/show/hide/synchronize the pane
- `src/OfficeAgent.DesktopUI/Views/ChatPaneView.xaml`
  Purpose: main chat UI
- `src/OfficeAgent.DesktopUI/ViewModels/ChatPaneViewModel.cs`
  Purpose: message flow, selection display, confirm/cancel commands
- `src/OfficeAgent.DesktopUI/ViewModels/SettingsViewModel.cs`
  Purpose: API key, Base URL, model settings
- `src/OfficeAgent.Core/Models/*.cs`
  Purpose: session, message, selection, command, skill models
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
- `installer/OfficeAgent.Setup`
  Purpose: MSI packaging project
- `tests/OfficeAgent.Core.Tests/*`
  Purpose: logic, routing, confirmation, payload, persistence tests
- `tests/OfficeAgent.Infrastructure.Tests/*`
  Purpose: file storage and HTTP tests
- `docs/vsto-manual-test-checklist.md`
  Purpose: Excel desktop + installer verification checklist

## Task 1: Create the VSTO Solution Skeleton

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn`
- Create: `src/OfficeAgent.DesktopUI`
- Create: `src/OfficeAgent.Core`
- Create: `src/OfficeAgent.Infrastructure`
- Create: `tests/OfficeAgent.Core.Tests`

- [ ] Create a new Excel VSTO Add-in solution targeting `.NET Framework 4.8`.
- [ ] Add three class library projects: `OfficeAgent.Core`, `OfficeAgent.Infrastructure`, `OfficeAgent.DesktopUI`, all targeting `net48`.
- [ ] Add a test project for pure logic and storage code.
- [ ] Add project references:
  - `OfficeAgent.ExcelAddIn` -> `OfficeAgent.Core`, `OfficeAgent.Infrastructure`, `OfficeAgent.DesktopUI`
  - `OfficeAgent.Infrastructure` -> `OfficeAgent.Core`
  - `OfficeAgent.DesktopUI` -> `OfficeAgent.Core`
- [ ] Verify F5 can launch Excel with the empty add-in loaded.
- [ ] Commit with message: `chore: scaffold vsto officeagent solution`

## Task 2: Add Ribbon and Custom Task Pane Host

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Create: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- Create: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

- [ ] Add a Ribbon tab/group/button named `OfficeAgent`.
- [ ] Create `TaskPaneHostControl` as a WinForms `UserControl`.
- [ ] In `ThisAddIn_Startup`, initialize a `TaskPaneController`.
- [ ] Make the Ribbon button show/hide the custom task pane.
- [ ] Set a stable width and default dock position on the right.
- [ ] Verify Excel can repeatedly open/close the pane without duplicate instances.
- [ ] Commit with message: `feat: add vsto ribbon and task pane shell`

## Task 3: Build the Chat UI Shell in WPF

**Files:**
- Create: `src/OfficeAgent.DesktopUI/Views/ChatPaneView.xaml`
- Create: `src/OfficeAgent.DesktopUI/ViewModels/ChatPaneViewModel.cs`
- Create: `src/OfficeAgent.DesktopUI/ViewModels/MessageItemViewModel.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`

- [ ] Host a WPF `ChatPaneView` inside the WinForms task pane using `ElementHost`.
- [ ] Add UI regions for:
  - session list
  - message thread
  - composer
  - selection badge
  - settings panel
- [ ] Implement a basic ViewModel with mock welcome message and send command.
- [ ] Keep the UI state local first; no Excel or API calls in this task.
- [ ] Verify task pane resizing and scroll behavior inside Excel.
- [ ] Commit with message: `feat: add desktop chat pane ui shell`

## Task 4: Add Local Session and Settings Persistence

**Files:**
- Create: `src/OfficeAgent.Core/Models/ChatSession.cs`
- Create: `src/OfficeAgent.Core/Models/ChatMessage.cs`
- Create: `src/OfficeAgent.Core/Models/AppSettings.cs`
- Create: `src/OfficeAgent.Core/Services/ISessionStore.cs`
- Create: `src/OfficeAgent.Core/Services/ISettingsStore.cs`
- Create: `src/OfficeAgent.Infrastructure/Storage/FileSessionStore.cs`
- Create: `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`
- Create: `src/OfficeAgent.Infrastructure/Security/DpapiSecretProtector.cs`
- Modify: `src/OfficeAgent.DesktopUI/ViewModels/ChatPaneViewModel.cs`
- Modify: `src/OfficeAgent.DesktopUI/ViewModels/SettingsViewModel.cs`
- Test: `tests/OfficeAgent.Core.Tests/SessionStoreTests.cs`

- [ ] Store sessions under `%LocalAppData%\OfficeAgent\sessions\`.
- [ ] Store non-sensitive settings in `%LocalAppData%\OfficeAgent\settings.json`.
- [ ] Encrypt `API Key` with user-scope DPAPI before persistence.
- [ ] Support:
  - create session
  - switch session
  - delete session
  - reopen last active session
- [ ] Add `Base URL` to settings from day one.
- [ ] Add tests for serialization, malformed file recovery, and encrypted secret roundtrip.
- [ ] Commit with message: `feat: persist sessions and settings for vsto add-in`

## Task 5: Add Excel Selection Context and Event Bridge

**Files:**
- Create: `src/OfficeAgent.Core/Models/SelectionContext.cs`
- Create: `src/OfficeAgent.Core/Services/IExcelContextService.cs`
- Create: `src/OfficeAgent.Infrastructure/Excel/ExcelSelectionContextService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.DesktopUI/ViewModels/ChatPaneViewModel.cs`
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
- [ ] Push updates to the WPF ViewModel on the UI thread.
- [ ] Block unsupported multi-area selections with a clear user message.
- [ ] Add tests for normalization logic independent of Excel runtime.
- [ ] Commit with message: `feat: add excel selection context bridge`

## Task 6: Add Excel Command Execution and Confirmation Policy

**Files:**
- Create: `src/OfficeAgent.Core/Models/ExcelCommand.cs`
- Create: `src/OfficeAgent.Core/Services/IExcelCommandExecutor.cs`
- Create: `src/OfficeAgent.Core/Services/ConfirmationService.cs`
- Create: `src/OfficeAgent.Infrastructure/Excel/ExcelInteropAdapter.cs`
- Modify: `src/OfficeAgent.DesktopUI/ViewModels/ChatPaneViewModel.cs`
- Test: `tests/OfficeAgent.Core.Tests/ConfirmationServiceTests.cs`

- [ ] Implement a command classification policy:
  - `read` commands execute immediately
  - `write` commands require confirmation
- [ ] Add executor methods for:
  - read current selection as table
  - write range
  - add/rename/delete worksheet
- [ ] Show a confirmation card in the UI before write execution.
- [ ] Keep executor behind an interface so the ViewModel never touches Interop objects directly.
- [ ] Add tests for confirmation rules and command model validation.
- [ ] Commit with message: `feat: add excel command executor and confirmation flow`

## Task 7: Add Agent Orchestration and Skill Routing

**Files:**
- Create: `src/OfficeAgent.Core/Models/AgentCommandEnvelope.cs`
- Create: `src/OfficeAgent.Core/Services/IAgentOrchestrator.cs`
- Create: `src/OfficeAgent.Core/Services/ISkillRegistry.cs`
- Create: `src/OfficeAgent.Core/Orchestration/AgentOrchestrator.cs`
- Create: `src/OfficeAgent.Core/Skills/SkillRegistry.cs`
- Test: `tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs`

- [ ] Preserve the current product contract:
  - natural language first
  - slash command as forced skill entry
  - structured command envelope instead of arbitrary script text
- [ ] Support three routes:
  - chat
  - excel command
  - skill
- [ ] Add `upload_data` route detection from both `/upload_data ...` and natural language.
- [ ] Add tests for route resolution and invalid command rejection.
- [ ] Commit with message: `feat: add vsto agent orchestration and skill routing`

## Task 8: Implement HTTP Clients and upload_data Skill

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Http/LlmClient.cs`
- Create: `src/OfficeAgent.Infrastructure/Http/BusinessApiClient.cs`
- Create: `src/OfficeAgent.Core/Skills/UploadDataSkill.cs`
- Create: `src/OfficeAgent.Core/Models/UploadPreview.cs`
- Modify: `src/OfficeAgent.DesktopUI/ViewModels/ChatPaneViewModel.cs`
- Test: `tests/OfficeAgent.Core.Tests/UploadDataSkillTests.cs`

- [ ] Use `HttpClient` with configurable `Base URL`.
- [ ] Read API Key from the protected settings store.
- [ ] Build the `upload_data` flow:
  - read selection table
  - infer columns
  - build preview payload
  - show confirmation
  - submit to `upload_data_api`
  - render result in the chat thread
- [ ] Add retry/timeout/error formatting at the client layer.
- [ ] Add tests for payload building, base URL normalization, and API failure handling.
- [ ] Commit with message: `feat: add upload_data skill for vsto add-in`

## Task 9: Package and Validate Enterprise Distribution

**Files:**
- Create: `installer/OfficeAgent.Setup`
- Create: `docs/vsto-manual-test-checklist.md`
- Modify: solution packaging settings

- [ ] Build an MSI installer that deploys the add-in for target users or machines.
- [ ] Include prerequisites documentation:
  - VSTO runtime
  - .NET Framework 4.8
  - Office compatibility expectations
- [ ] Add installer options for:
  - install
  - uninstall
  - upgrade
- [ ] Verify add-in load behavior after fresh install and after Excel restart.
- [ ] Verify no manual sideload, manifest upload, or localhost trust is required for end users.
- [ ] Write a manual checklist for:
  - Excel 2019 x86/x64
  - selection updates
  - upload_data confirmation
  - local session restore
  - API Key/Base URL save and reload
- [ ] Commit with message: `chore: package vsto add-in for enterprise deployment`

## Task 10: Stabilization and Release Readiness

**Files:**
- Modify: logging, diagnostics, settings recovery, docs as needed

- [ ] Add structured logging for startup, pane open/close, selection changes, skill execution, and HTTP failures.
- [ ] Add defensive handling for:
  - protected worksheets
  - merged cells
  - invalid selection
  - workbook closed during pending action
- [ ] Run the full test suite.
- [ ] Run manual QA in supported Excel environments.
- [ ] Produce a release note and known-limitations document.
- [ ] Commit with message: `chore: harden vsto officeagent mvp`

