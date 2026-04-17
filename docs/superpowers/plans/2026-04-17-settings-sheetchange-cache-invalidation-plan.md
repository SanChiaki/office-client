# Settings SheetChange Cache Invalidation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Automatically invalidate `_Settings` metadata caches when users manually edit the `_Settings` worksheet.

**Architecture:** Keep the existing `_Settings` cache and bulk-read optimizations, but add an explicit cache invalidation hook on the metadata store and call it from an Excel `SheetChange` handler only when the changed sheet is `_Settings`. Reset ribbon refresh state at the same time so the next business-sheet interaction reloads fresh binding data.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, xUnit

---

## File Structure

- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
  Responsibility: expose an internal cache invalidation entry point and clear both bindings and field-mapping caches.
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Responsibility: expose an internal refresh-state invalidation entry point so the next same-sheet refresh is not skipped.
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Responsibility: subscribe to `Application.SheetChange`, detect `_Settings`, invalidate metadata cache, and reset project refresh state.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`
  Responsibility: verify invalidation clears cached table reads.
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Responsibility: verify invalidating refresh state forces a reload on the same active sheet.
- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  Responsibility: verify source wiring for `_Settings` `SheetChange` invalidation exists in `ThisAddIn`.

### Task 1: Add Failing Tests

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Write the failing tests**
- [ ] **Step 2: Run `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetMetadataStoreTests|FullyQualifiedName~RibbonSyncControllerTests|FullyQualifiedName~AgentRibbonConfigurationTests"` and verify the new tests fail**

### Task 2: Implement SheetChange-Based Invalidation

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

- [ ] **Step 1: Add minimal invalidation methods**
- [ ] **Step 2: Wire `Application.SheetChange` to `_Settings` cache invalidation**
- [ ] **Step 3: Re-run the targeted test filter and verify it passes**

### Task 3: Full Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`

- [ ] **Step 1: Document automatic `_Settings` edit invalidation**
- [ ] **Step 2: Run `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --no-restore`**
- [ ] **Step 3: Run `pwsh -NoProfile -ExecutionPolicy Bypass -File eng\Build-VstoAddIn.ps1 -ProjectPath src\OfficeAgent.ExcelAddIn\OfficeAgent.ExcelAddIn.csproj -Configuration Debug`**
