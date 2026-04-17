# Settings Metadata Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce data-sheet interaction lag caused by large `_Settings` metadata tables without changing the visible metadata layout.

**Architecture:** Stop refreshing active project metadata on every selection move and instead refresh only when the active worksheet changes. Keep the `_Settings` sheet layout intact, but make metadata reads cheaper by loading `UsedRange.Value2` in bulk and caching parsed bindings and field mappings within the metadata store and execution flow. Replace repeated linear mapping scans with per-layout lookup indexes so download/upload preparation cost scales with actual headers, not full mapping row count.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, xUnit

---

## File Structure

- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Responsibility: control when ribbon project state refreshes in response to Excel events.
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Responsibility: expose a safe refresh operation that can skip redundant metadata reloads for the same active worksheet.
- `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs`
  Responsibility: read `_Settings` with bulk `UsedRange.Value2` access instead of cell-by-cell COM calls.
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
  Responsibility: cache parsed `SheetBinding` and `SheetFieldMappings` rows and invalidate cache on writes.
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs`
  Responsibility: pre-index mapping rows for single-header and activity-property lookup.
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Responsibility: cover redundant refresh avoidance behavior.
- `tests/OfficeAgent.ExcelAddIn.Tests/ExcelWorkbookMetadataAdapterTests.cs`
  Responsibility: cover bulk metadata reads and preserve current parsing semantics.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
  Responsibility: cover mapping reuse/indexed matching without changing upload/download behavior.

### Task 1: Throttle Active Project Refreshes to Sheet Changes

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`

- [ ] **Step 1: Write the failing regression tests**

Add tests proving repeated refreshes for the same active sheet do not reload metadata, while switching sheet names still reloads and updates state.

- [ ] **Step 2: Run the targeted tests to verify they fail**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests`
Expected: FAIL on the new redundant-refresh regression tests.

- [ ] **Step 3: Implement minimal sheet-aware refresh throttling**

Track the last refreshed sheet name, refresh on startup and true sheet change, and keep selection-context publishing intact.

- [ ] **Step 4: Re-run the targeted tests**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests`
Expected: PASS for all `RibbonSyncControllerTests`.

### Task 2: Replace Cell-by-Cell Metadata Reads with Bulk Value2 Parsing

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/ExcelWorkbookMetadataAdapterTests.cs`

- [ ] **Step 1: Write the failing adapter tests**

Add tests that exercise `_Settings` reads through a fake used range returning a 2D `Value2` payload, including single-cell and multi-cell sections.

- [ ] **Step 2: Run the targeted tests to verify they fail**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~ExcelWorkbookMetadataAdapterTests`
Expected: FAIL on the new bulk-read tests.

- [ ] **Step 3: Implement minimal bulk parsing**

Replace per-cell COM iteration with `UsedRange.Value2`, normalize scalar versus `object[,]` shapes, preserve blank-row trimming, and keep active-sheet restoration behavior unchanged.

- [ ] **Step 4: Re-run the targeted tests**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~ExcelWorkbookMetadataAdapterTests`
Expected: PASS for all `ExcelWorkbookMetadataAdapterTests`.

### Task 3: Cache Bindings and Field Mappings, Then Index Header Matching

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] **Step 1: Write the failing execution-service tests**

Add tests covering repeated metadata loads for the same sheet and large mapping sets, plus matching behavior for both single-row and two-row headers after indexing.

- [ ] **Step 2: Run the targeted tests to verify they fail**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSyncExecutionServiceTests`
Expected: FAIL on the new caching/indexing regression tests.

- [ ] **Step 3: Implement minimal caching and indexes**

Cache `SheetBinding` and `SheetFieldMappings` by sheet identity, invalidate caches on save/clear, and build header lookup dictionaries once per match operation instead of scanning every mapping row for every worksheet column.

- [ ] **Step 4: Re-run the targeted tests**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSyncExecutionServiceTests`
Expected: PASS for all `WorksheetSyncExecutionServiceTests`.

### Task 4: Full Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`

- [ ] **Step 1: Update behavior documentation**

Document that project metadata refresh now tracks sheet changes rather than every selection move, and note that `_Settings` reads are now bulk-read/cached for large metadata tables.

- [ ] **Step 2: Run the Excel add-in test assembly**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --no-restore`
Expected: PASS with `0` failures.

- [ ] **Step 3: Run the signed VSTO build**

Run: `pwsh -NoProfile -ExecutionPolicy Bypass -File eng\Build-VstoAddIn.ps1 -ProjectPath src\OfficeAgent.ExcelAddIn\OfficeAgent.ExcelAddIn.csproj -Configuration Debug`
Expected: `0` warnings and `0` errors.

- [ ] **Step 4: Commit**

```bash
git add docs/modules/ribbon-sync-current-behavior.md docs/superpowers/plans/2026-04-17-settings-metadata-performance-implementation-plan.md src/OfficeAgent.ExcelAddIn/ThisAddIn.cs src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs tests/OfficeAgent.ExcelAddIn.Tests/ExcelWorkbookMetadataAdapterTests.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs
git commit -m "perf: reduce settings metadata read overhead"
```
