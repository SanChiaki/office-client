# Ribbon Sync Excel I/O Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce Ribbon Sync Excel-side latency by batching full-download writes, batching full-upload reads, caching repeated row-ID lookups for partial sync, and wrapping heavy Excel operations in a safe bulk-operation scope.

**Architecture:** Keep the current Ribbon Sync workflow and connector contracts intact, but extend the worksheet grid adapter with bulk capabilities. Use a small helper to segment non-contiguous managed columns for download writes, a dedicated upload-value normalizer plus safe per-cell fallback for risky formats, and an adapter-owned bulk-operation scope so Excel global state management stays out of business orchestration.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, xUnit

---

## File Structure

- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegment.cs`
  Responsibility: represent one contiguous managed-column segment for batch writes.
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegmentBuilder.cs`
  Responsibility: split runtime columns into contiguous write segments.
- Create: `src/OfficeAgent.ExcelAddIn/Excel/ExcelUploadValueNormalizer.cs`
  Responsibility: normalize bulk-read Excel values into upload strings and signal when per-cell `Text` fallback is required.
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`
  Responsibility: expose bulk read/write methods and a bulk-operation scope.
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
  Responsibility: implement `Range.Value2` / `Range.NumberFormat` bulk access plus Excel state scope handling.
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
  Responsibility: orchestrate segmented batch writes, bulk upload reads, normalizer fallback, and cached row-ID lookup.
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
  Responsibility: regression coverage for full download/write batching, full upload/read batching, fallback reads, row-ID caching, and bulk-operation scope usage.
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetColumnSegmentBuilderTests.cs`
  Responsibility: unit-test contiguous managed-column segmentation.
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/ExcelUploadValueNormalizerTests.cs`
  Responsibility: unit-test upload value normalization and unsafe-format fallback detection.
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: document the new batch read/write behavior and the remaining fallback boundary.

### Task 1: Add Column Segment Builder with TDD

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetColumnSegmentBuilderTests.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegment.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegmentBuilder.cs`

- [ ] **Step 1: Write the failing segment-builder tests**

```csharp
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Excel;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetColumnSegmentBuilderTests
    {
        [Fact]
        public void BuildGroupsContiguousManagedColumnsIntoSegments()
        {
            var builder = new WorksheetColumnSegmentBuilder();
            var segments = builder.Build(new[]
            {
                new WorksheetRuntimeColumn { ColumnIndex = 1, ApiFieldKey = "row_id", IsIdColumn = true },
                new WorksheetRuntimeColumn { ColumnIndex = 2, ApiFieldKey = "owner_name" },
                new WorksheetRuntimeColumn { ColumnIndex = 4, ApiFieldKey = "start_12345678" },
                new WorksheetRuntimeColumn { ColumnIndex = 5, ApiFieldKey = "end_12345678" },
            });

            Assert.Collection(
                segments,
                first =>
                {
                    Assert.Equal(1, first.StartColumn);
                    Assert.Equal(2, first.EndColumn);
                    Assert.Equal(new[] { 1, 2 }, first.Columns.Select(column => column.ColumnIndex).ToArray());
                },
                second =>
                {
                    Assert.Equal(4, second.StartColumn);
                    Assert.Equal(5, second.EndColumn);
                    Assert.Equal(new[] { 4, 5 }, second.Columns.Select(column => column.ColumnIndex).ToArray());
                });
        }

        [Fact]
        public void BuildSkipsNullColumnsAndReturnsEmptyForNoManagedColumns()
        {
            var builder = new WorksheetColumnSegmentBuilder();
            var segments = builder.Build(new WorksheetRuntimeColumn[] { null });

            Assert.Empty(segments);
        }
    }
}
```

- [ ] **Step 2: Run the new tests to verify they fail**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetColumnSegmentBuilderTests
```

Expected: FAIL because `WorksheetColumnSegmentBuilder` and `WorksheetColumnSegment` do not exist yet.

- [ ] **Step 3: Implement the segment model and builder**

`src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegment.cs`

```csharp
using System;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetColumnSegment
    {
        public int StartColumn { get; set; }

        public int EndColumn { get; set; }

        public WorksheetRuntimeColumn[] Columns { get; set; } = Array.Empty<WorksheetRuntimeColumn>();
    }
}
```

`src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegmentBuilder.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetColumnSegmentBuilder
    {
        public WorksheetColumnSegment[] Build(IReadOnlyList<WorksheetRuntimeColumn> columns)
        {
            var ordered = (columns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null)
                .OrderBy(column => column.ColumnIndex)
                .ToArray();
            if (ordered.Length == 0)
            {
                return Array.Empty<WorksheetColumnSegment>();
            }

            var segments = new List<WorksheetColumnSegment>();
            var currentColumns = new List<WorksheetRuntimeColumn> { ordered[0] };
            var currentStart = ordered[0].ColumnIndex;
            var currentEnd = ordered[0].ColumnIndex;

            for (var index = 1; index < ordered.Length; index++)
            {
                var column = ordered[index];
                if (column.ColumnIndex == currentEnd + 1)
                {
                    currentColumns.Add(column);
                    currentEnd = column.ColumnIndex;
                    continue;
                }

                segments.Add(new WorksheetColumnSegment
                {
                    StartColumn = currentStart,
                    EndColumn = currentEnd,
                    Columns = currentColumns.ToArray(),
                });

                currentColumns = new List<WorksheetRuntimeColumn> { column };
                currentStart = column.ColumnIndex;
                currentEnd = column.ColumnIndex;
            }

            segments.Add(new WorksheetColumnSegment
            {
                StartColumn = currentStart,
                EndColumn = currentEnd,
                Columns = currentColumns.ToArray(),
            });

            return segments.ToArray();
        }
    }
}
```

- [ ] **Step 4: Run the segment-builder tests to verify they pass**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetColumnSegmentBuilderTests
```

Expected: PASS with `2` tests passed.

- [ ] **Step 5: Commit the segment-builder slice**

```powershell
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetColumnSegmentBuilderTests.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegment.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetColumnSegmentBuilder.cs
git commit -m "test: add managed column segment builder"
```

### Task 2: Add Upload Value Normalizer with Safe Fallback Rules

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/ExcelUploadValueNormalizerTests.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/ExcelUploadValueNormalizer.cs`

- [ ] **Step 1: Write the failing normalizer tests**

```csharp
using OfficeAgent.ExcelAddIn.Excel;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelUploadValueNormalizerTests
    {
        [Fact]
        public void TryNormalizeReturnsPlainTextForSafeGeneralValues()
        {
            var normalizer = new ExcelUploadValueNormalizer();

            Assert.True(normalizer.TryNormalize(null, "General", out var emptyText));
            Assert.Equal(string.Empty, emptyText);

            Assert.True(normalizer.TryNormalize("Alpha", "@", out var stringText));
            Assert.Equal("Alpha", stringText);

            Assert.True(normalizer.TryNormalize(12d, "General", out var integerText));
            Assert.Equal("12", integerText);

            Assert.True(normalizer.TryNormalize(12.5d, "General", out var decimalText));
            Assert.Equal("12.5", decimalText);
        }

        [Fact]
        public void TryNormalizeReturnsFalseForDateAndFormattedNumericCells()
        {
            var normalizer = new ExcelUploadValueNormalizer();

            Assert.False(normalizer.TryNormalize(46024d, "yyyy/m/d", out _));
            Assert.False(normalizer.TryNormalize(0.25d, "0%", out _));
            Assert.False(normalizer.TryNormalize(123d, "000000", out _));
        }
    }
}
```

- [ ] **Step 2: Run the normalizer tests to verify they fail**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~ExcelUploadValueNormalizerTests
```

Expected: FAIL because `ExcelUploadValueNormalizer` does not exist yet.

- [ ] **Step 3: Implement the normalizer**

`src/OfficeAgent.ExcelAddIn/Excel/ExcelUploadValueNormalizer.cs`

```csharp
using System;
using System.Globalization;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelUploadValueNormalizer
    {
        public bool TryNormalize(object value, string numberFormat, out string normalized)
        {
            if (value == null)
            {
                normalized = string.Empty;
                return true;
            }

            if (value is string text)
            {
                normalized = text;
                return true;
            }

            if (value is bool booleanValue)
            {
                normalized = Convert.ToString(booleanValue, CultureInfo.InvariantCulture) ?? string.Empty;
                return true;
            }

            if (value is double numericValue)
            {
                if (RequiresDisplayTextFallback(numberFormat))
                {
                    normalized = null;
                    return false;
                }

                normalized = numericValue % 1d == 0d
                    ? numericValue.ToString("0", CultureInfo.InvariantCulture)
                    : numericValue.ToString("0.###############", CultureInfo.InvariantCulture);
                return true;
            }

            normalized = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return true;
        }

        private static bool RequiresDisplayTextFallback(string numberFormat)
        {
            var format = numberFormat ?? string.Empty;
            if (string.IsNullOrWhiteSpace(format) || string.Equals(format, "General", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return format.Contains("%") ||
                   format.Contains("y") ||
                   format.Contains("m") ||
                   format.Contains("d") ||
                   format.Contains("h") ||
                   format.Contains("s") ||
                   format.Contains("0");
        }
    }
}
```

Note: the follow-up implementation task will tighten `RequiresDisplayTextFallback()` if the first pass is too aggressive or too loose; this slice exists to lock in the fallback contract before orchestration work starts.

- [ ] **Step 4: Run the normalizer tests to verify they pass**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~ExcelUploadValueNormalizerTests
```

Expected: PASS with `2` tests passed.

- [ ] **Step 5: Commit the normalizer slice**

```powershell
git add tests/OfficeAgent.ExcelAddIn.Tests/ExcelUploadValueNormalizerTests.cs src/OfficeAgent.ExcelAddIn/Excel/ExcelUploadValueNormalizer.cs
git commit -m "test: add excel upload value normalizer"
```

### Task 3: Add Bulk Grid APIs and Batch Full-Download Writes

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] **Step 1: Add failing full-download batching tests**

Append these tests to `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`:

```csharp
[Fact]
public void ExecuteFullDownloadUsesBatchWriteForContiguousManagedColumns()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    SeedRecognizedHeaders(grid, "Sheet1", binding);

    var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.Single(grid.WriteRangeCalls);
    Assert.Equal(6, grid.WriteRangeCalls[0].StartRow);
    Assert.Equal(1, grid.WriteRangeCalls[0].StartColumn);
    Assert.Equal(1, grid.WriteRangeCalls[0].RowCount);
    Assert.Equal(4, grid.WriteRangeCalls[0].ColumnCount);
}

[Fact]
public void ExecuteFullDownloadSplitsBatchWritesAcrossNonContiguousManagedColumns()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    grid.SetCell("Sheet1", 3, 1, "ID");
    grid.SetCell("Sheet1", 3, 2, "项目负责人");
    grid.SetCell("Sheet1", 3, 4, "测试活动111");
    grid.SetCell("Sheet1", 4, 4, "开始时间");
    grid.SetCell("Sheet1", 4, 5, "结束时间");
    grid.SetCell("Sheet1", 6, 3, "用户备注");

    var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.Equal(2, grid.WriteRangeCalls.Count);
    Assert.Equal("用户备注", grid.GetCell("Sheet1", 6, 3));
}
```

Also extend `FakeWorksheetGridAdapter` with call tracking fields so the tests compile once the interface is updated:

```csharp
public List<WriteRangeRecord> WriteRangeCalls { get; } = new List<WriteRangeRecord>();
public int BeginBulkOperationCount { get; private set; }
public int EndBulkOperationCount { get; private set; }
```

- [ ] **Step 2: Run the full-download batching tests to verify they fail**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecuteFullDownloadUsesBatchWriteForContiguousManagedColumns|FullyQualifiedName~ExecuteFullDownloadSplitsBatchWritesAcrossNonContiguousManagedColumns"
```

Expected: FAIL because `IWorksheetGridAdapter` and the fake grid do not support bulk writes yet.

- [ ] **Step 3: Extend the grid adapter interface and implement bulk writes**

`src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IWorksheetGridAdapter
    {
        string GetCellText(string sheetName, int row, int column);
        void SetCellText(string sheetName, int row, int column, string value);
        void WriteRangeValues(string sheetName, int startRow, int startColumn, object[,] values);
        object[,] ReadRangeValues(string sheetName, int startRow, int endRow, int startColumn, int endColumn);
        string[,] ReadRangeNumberFormats(string sheetName, int startRow, int endRow, int startColumn, int endColumn);
        IDisposable BeginBulkOperation();
        void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn);
        void ClearWorksheet(string sheetName);
        void MergeCells(string sheetName, int row, int column, int rowSpan, int columnSpan);
        int GetLastUsedRow(string sheetName);
        int GetLastUsedColumn(string sheetName);
    }
}
```

Add the write-side implementation to `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`:

```csharp
public void WriteRangeValues(string sheetName, int startRow, int startColumn, object[,] values)
{
    if (values == null)
    {
        throw new ArgumentNullException(nameof(values));
    }

    var worksheet = GetWorksheet(sheetName);
    var rowCount = values.GetLength(0);
    var columnCount = values.GetLength(1);
    if (rowCount == 0 || columnCount == 0)
    {
        return;
    }

    var topLeft = worksheet.Cells[startRow, startColumn] as ExcelInterop.Range;
    var writeTarget = topLeft.get_Resize(rowCount, columnCount);
    writeTarget.Value2 = values;
}
```

Add a first-pass bulk-operation scope to the same class so later tasks can start using it without changing signatures again:

```csharp
public IDisposable BeginBulkOperation()
{
    return new BulkOperationScope(application);
}
```

Update `FakeWorksheetGridAdapter` in `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs` to handle the new methods and record calls:

```csharp
case "WriteRangeValues":
    WriteRange(
        (string)call.InArgs[0],
        (int)call.InArgs[1],
        (int)call.InArgs[2],
        (object[,])call.InArgs[3]);
    return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
case "BeginBulkOperation":
    BeginBulkOperationCount++;
    return new ReturnMessage(new DelegateDisposable(() => EndBulkOperationCount++), null, 0, call.LogicalCallContext, call);
```

- [ ] **Step 4: Refactor full download to use contiguous column segments and batch writes**

Update the relevant parts of `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`:

```csharp
private readonly WorksheetColumnSegmentBuilder segmentBuilder;

public WorksheetSyncExecutionService(
    WorksheetSyncService worksheetSyncService,
    IWorksheetMetadataStore metadataStore,
    IWorksheetSelectionReader selectionReader,
    IWorksheetGridAdapter gridAdapter,
    SyncOperationPreviewFactory previewFactory)
{
    this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
    _ = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
    this.selectionReader = selectionReader ?? throw new ArgumentNullException(nameof(selectionReader));
    this.gridAdapter = gridAdapter ?? throw new ArgumentNullException(nameof(gridAdapter));
    this.previewFactory = previewFactory ?? throw new ArgumentNullException(nameof(previewFactory));
    selectionResolver = new WorksheetSelectionResolver();
    layoutService = new WorksheetSchemaLayoutService();
    valueAccessor = new FieldMappingValueAccessor();
    headerMatcher = new WorksheetHeaderMatcher(valueAccessor);
    segmentBuilder = new WorksheetColumnSegmentBuilder();
}

private void WriteFullWorksheet(WorksheetDownloadPlan plan)
{
    var binding = plan.Binding;
    var columns = plan.RuntimeColumns ?? Array.Empty<WorksheetRuntimeColumn>();
    var clearEndRow = Math.Max(gridAdapter.GetLastUsedRow(plan.SheetName), binding.DataStartRow + (plan.Rows?.Count ?? 0) + 10);

    using (gridAdapter.BeginBulkOperation())
    {
        ClearManagedArea(plan.SheetName, binding, columns, plan.UsesExistingLayout, clearEndRow);

        if (!plan.UsesExistingLayout)
        {
            var headerPlan = layoutService.BuildHeaderPlan(binding, columns);
            foreach (var headerCell in headerPlan)
            {
                gridAdapter.SetCellText(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.Text);
                gridAdapter.MergeCells(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.RowSpan, headerCell.ColumnSpan);
            }
        }

        foreach (var segment in segmentBuilder.Build(columns))
        {
            var values = BuildSegmentValues(segment, plan.Rows);
            gridAdapter.WriteRangeValues(plan.SheetName, binding.DataStartRow, segment.StartColumn, values);
        }
    }
}
```

- [ ] **Step 5: Run the targeted batching tests to verify they pass**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecuteFullDownloadUsesBatchWriteForContiguousManagedColumns|FullyQualifiedName~ExecuteFullDownloadSplitsBatchWritesAcrossNonContiguousManagedColumns|FullyQualifiedName~ExecuteFullDownloadDoesNotRewriteExistingRecognizedHeaders"
```

Expected: PASS with the two new tests plus the existing recognized-header regression.

- [ ] **Step 6: Commit the full-download batching slice**

```powershell
git add src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs
git commit -m "perf: batch ribbon sync full download writes"
```

### Task 4: Batch Full-Upload Reads and Cache Partial Row IDs

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] **Step 1: Add failing upload-batching and row-ID cache tests**

Append these tests to `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`:

```csharp
[Fact]
public void ExecuteFullUploadUsesBatchReadForManagedRegion()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    SeedRecognizedHeaders(grid, "Sheet1", binding);
    grid.SetCell("Sheet1", 6, 1, "row-1");
    grid.SetCell("Sheet1", 6, 2, "李四");
    grid.SetCell("Sheet1", 6, 3, "2026-01-02");
    grid.SetCell("Sheet1", 6, 4, "2026-01-05");

    var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
    var preview = ReadPreview(plan);

    Assert.Equal(3, preview.Changes.Length);
    Assert.Single(grid.ReadRangeCalls);
    Assert.Equal(6, grid.ReadRangeCalls[0].StartRow);
    Assert.Equal(1, grid.ReadRangeCalls[0].StartColumn);
    Assert.Equal(4, grid.ReadRangeCalls[0].EndColumn);
}

[Fact]
public void ExecuteFullUploadFallsBackToCellTextForUnsafeFormattedCells()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    SeedRecognizedHeaders(grid, "Sheet1", binding);
    grid.SetRawCell("Sheet1", 6, 1, "row-1", "@");
    grid.SetRawCell("Sheet1", 6, 2, "李四", "@");
    grid.SetRawCell("Sheet1", 6, 3, 46024d, "yyyy/m/d");
    grid.SetCell("Sheet1", 6, 3, "2026/1/2");
    grid.SetRawCell("Sheet1", 6, 4, "2026-01-05", "@");

    var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
    var preview = ReadPreview(plan);

    Assert.Contains(preview.Changes, change => change.ApiFieldKey == "start_12345678" && change.NewValue == "2026/1/2");
    Assert.Contains(grid.GetCellTextCalls, call => call.SheetName == "Sheet1" && call.Row == 6 && call.Column == 3);
}

[Fact]
public void PreparePartialUploadReadsEachRowIdAtMostOncePerRow()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

    var selectionReader = new FakeWorksheetSelectionReader
    {
        VisibleCells = new[]
        {
            new SelectedVisibleCell { Row = 6, Column = 2, Value = "李四" },
            new SelectedVisibleCell { Row = 6, Column = 3, Value = "2026-01-02" },
        },
    };
    var (service, grid) = CreateService(connector, metadataStore, selectionReader);
    SeedRecognizedHeaders(grid, "Sheet1", binding);
    grid.SetCell("Sheet1", 6, 1, "row-1");
    grid.SetCell("Sheet1", 6, 2, "李四");
    grid.SetCell("Sheet1", 6, 3, "2026-01-02");

    InvokePrepare(service, "PreparePartialUpload", "Sheet1");

    Assert.Equal(1, grid.GetCellTextCalls.Count(call => call.SheetName == "Sheet1" && call.Row == 6 && call.Column == 1));
}

[Fact]
public void ExecutePartialDownloadReadsEachRowIdAtMostOncePerRow()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    connector.FindResult = new[] { CreateRow("row-1", "李四", "2026-02-01", "2026-02-09") };

    var selectionReader = new FakeWorksheetSelectionReader
    {
        VisibleCells = new[]
        {
            new SelectedVisibleCell { Row = 6, Column = 2, Value = "李四" },
            new SelectedVisibleCell { Row = 6, Column = 3, Value = "旧开始时间" },
        },
    };
    var (service, grid) = CreateService(connector, metadataStore, selectionReader);
    SeedRecognizedHeaders(grid, "Sheet1", binding);
    grid.SetCell("Sheet1", 6, 1, "row-1");
    grid.SetCell("Sheet1", 6, 2, "李四");
    grid.SetCell("Sheet1", 6, 3, "旧开始时间");

    var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.Equal(1, grid.GetCellTextCalls.Count(call => call.SheetName == "Sheet1" && call.Row == 6 && call.Column == 1));
}
```

- [ ] **Step 2: Run the new upload tests to verify they fail**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecuteFullUploadUsesBatchReadForManagedRegion|FullyQualifiedName~ExecuteFullUploadFallsBackToCellTextForUnsafeFormattedCells|FullyQualifiedName~PreparePartialUploadReadsEachRowIdAtMostOncePerRow|FullyQualifiedName~ExecutePartialDownloadReadsEachRowIdAtMostOncePerRow"
```

Expected: FAIL because the service still reads per-cell and the fake grid does not provide raw-value / number-format matrices yet.

- [ ] **Step 3: Implement bulk reads, normalization, and cached row-ID lookup**

Extend `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs` with bulk read helpers:

```csharp
public object[,] ReadRangeValues(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
{
    var worksheet = GetWorksheet(sheetName);
    var range = worksheet.Range[
        worksheet.Cells[startRow, startColumn],
        worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
    return ToTwoDimensionalObjectArray(range?.Value2, endRow - startRow + 1, endColumn - startColumn + 1);
}

public string[,] ReadRangeNumberFormats(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
{
    var worksheet = GetWorksheet(sheetName);
    var range = worksheet.Range[
        worksheet.Cells[startRow, startColumn],
        worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
    return ToTwoDimensionalStringArray(range?.NumberFormat, endRow - startRow + 1, endColumn - startColumn + 1);
}
```

Refactor `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`:

```csharp
private readonly ExcelUploadValueNormalizer uploadValueNormalizer;

public WorksheetSyncExecutionService(
    WorksheetSyncService worksheetSyncService,
    IWorksheetMetadataStore metadataStore,
    IWorksheetSelectionReader selectionReader,
    IWorksheetGridAdapter gridAdapter,
    SyncOperationPreviewFactory previewFactory)
{
    this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
    _ = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
    this.selectionReader = selectionReader ?? throw new ArgumentNullException(nameof(selectionReader));
    this.gridAdapter = gridAdapter ?? throw new ArgumentNullException(nameof(gridAdapter));
    this.previewFactory = previewFactory ?? throw new ArgumentNullException(nameof(previewFactory));
    selectionResolver = new WorksheetSelectionResolver();
    layoutService = new WorksheetSchemaLayoutService();
    valueAccessor = new FieldMappingValueAccessor();
    headerMatcher = new WorksheetHeaderMatcher(valueAccessor);
    segmentBuilder = new WorksheetColumnSegmentBuilder();
    uploadValueNormalizer = new ExcelUploadValueNormalizer();
}

private Func<int, string> CreateCachedRowIdAccessor(string sheetName, WorksheetSchema schema)
{
    var cache = new Dictionary<int, string>();
    return row =>
    {
        if (!cache.TryGetValue(row, out var rowId))
        {
            rowId = GetRowId(sheetName, schema, row);
            cache[row] = rowId;
        }

        return rowId;
    };
}
```

Use `CreateCachedRowIdAccessor()` in:

- `ResolveCurrentSelection()`
- `WritePartialCells()`
- `ReadSelectionChanges()`

Then change full-upload reads to batch load `Value2` and `NumberFormat`, normalize safe cells, and only call `GetCellText()` when the normalizer returns `false`:

```csharp
var values = gridAdapter.ReadRangeValues(sheetName, binding.DataStartRow, lastUsedRow, startColumn, endColumn);
var formats = gridAdapter.ReadRangeNumberFormats(sheetName, binding.DataStartRow, lastUsedRow, startColumn, endColumn);

for (var rowOffset = 0; rowOffset < values.GetLength(0); rowOffset++)
{
    var absoluteRow = binding.DataStartRow + rowOffset;
    var rowIdRawValue = values[rowOffset, idColumn.ColumnIndex - startColumn];
    var rowIdFormat = formats[rowOffset, idColumn.ColumnIndex - startColumn];
    var rowId = uploadValueNormalizer.TryNormalize(rowIdRawValue, rowIdFormat, out var normalizedRowId)
        ? normalizedRowId
        : gridAdapter.GetCellText(sheetName, absoluteRow, idColumn.ColumnIndex);

    if (string.IsNullOrWhiteSpace(rowId))
    {
        continue;
    }

    foreach (var column in dataColumns.Where(item => !item.IsIdColumn))
    {
        var rawValue = values[rowOffset, column.ColumnIndex - startColumn];
        var numberFormat = formats[rowOffset, column.ColumnIndex - startColumn];
        var newValue = uploadValueNormalizer.TryNormalize(rawValue, numberFormat, out var normalized)
            ? normalized
            : gridAdapter.GetCellText(sheetName, absoluteRow, column.ColumnIndex);

        result.Add(new CellChange
        {
            SheetName = sheetName,
            RowId = rowId,
            ApiFieldKey = column.ApiFieldKey,
            OldValue = string.Empty,
            NewValue = newValue,
        });
    }
}
```

Update `FakeWorksheetGridAdapter` to support:

```csharp
public List<ReadRangeRecord> ReadRangeCalls { get; } = new List<ReadRangeRecord>();
public List<GetCellTextCall> GetCellTextCalls { get; } = new List<GetCellTextCall>();

case "ReadRangeValues":
    return new ReturnMessage(
        ReadRangeValues(
            (string)call.InArgs[0],
            (int)call.InArgs[1],
            (int)call.InArgs[2],
            (int)call.InArgs[3],
            (int)call.InArgs[4]),
        null,
        0,
        call.LogicalCallContext,
        call);
case "ReadRangeNumberFormats":
    return new ReturnMessage(
        ReadRangeNumberFormats(
            (string)call.InArgs[0],
            (int)call.InArgs[1],
            (int)call.InArgs[2],
            (int)call.InArgs[3],
            (int)call.InArgs[4]),
        null,
        0,
        call.LogicalCallContext,
        call);
```

- [ ] **Step 4: Run the upload and row-ID cache tests to verify they pass**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecuteFullUploadUsesBatchReadForManagedRegion|FullyQualifiedName~ExecuteFullUploadFallsBackToCellTextForUnsafeFormattedCells|FullyQualifiedName~PreparePartialUploadReadsEachRowIdAtMostOncePerRow|FullyQualifiedName~ExecutePartialDownloadReadsEachRowIdAtMostOncePerRow|FullyQualifiedName~ExecuteFullUploadUsesConfiguredDataStartRowAndRecognizedColumns|FullyQualifiedName~ExecutePartialDownloadUsesRecognizedHeadersAndIdLookupOutsideSelection"
```

Expected: PASS with the three new tests plus the existing full-upload regression.

- [ ] **Step 5: Commit the upload batching slice**

```powershell
git add src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs
git commit -m "perf: batch ribbon sync upload reads"
```

### Task 5: Tighten Bulk-Operation Scope, Update Docs, and Run Full Verification

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
- Modify: `docs/modules/ribbon-sync-current-behavior.md`

- [ ] **Step 1: Add failing tests for bulk-operation scope usage**

Append these tests to `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`:

```csharp
[Fact]
public void ExecuteFullDownloadBeginsAndEndsOneBulkOperation()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    SeedRecognizedHeaders(grid, "Sheet1", metadataStore.Bindings["Sheet1"]);

    var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.Equal(1, grid.BeginBulkOperationCount);
    Assert.Equal(1, grid.EndBulkOperationCount);
}

[Fact]
public void PrepareFullUploadBeginsAndEndsOneBulkOperation()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    SeedRecognizedHeaders(grid, "Sheet1", binding);
    grid.SetCell("Sheet1", 6, 1, "row-1");
    grid.SetCell("Sheet1", 6, 2, "李四");

    InvokePrepare(service, "PrepareFullUpload", "Sheet1");

    Assert.Equal(1, grid.BeginBulkOperationCount);
    Assert.Equal(1, grid.EndBulkOperationCount);
}
```

- [ ] **Step 2: Run the scope tests to verify they fail if any heavy path still escapes the scope**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecuteFullDownloadBeginsAndEndsOneBulkOperation|FullyQualifiedName~PrepareFullUploadBeginsAndEndsOneBulkOperation"
```

Expected: FAIL if any heavy path is still outside `using (gridAdapter.BeginBulkOperation())`.

- [ ] **Step 3: Finalize bulk-operation usage and update the behavior doc**

Ensure `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs` wraps:

- `WriteFullWorksheet()`
- `ReadAllCurrentCells()`

with:

```csharp
using (gridAdapter.BeginBulkOperation())
{
    // existing heavy read/write work
}
```

Tighten `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs` scope restore logic so it also preserves calculation mode:

```csharp
private sealed class BulkOperationScope : IDisposable
{
    private readonly ExcelInterop.Application application;
    private readonly bool previousScreenUpdating;
    private readonly bool previousEnableEvents;
    private readonly ExcelInterop.XlCalculation previousCalculation;

    public BulkOperationScope(ExcelInterop.Application application)
    {
        this.application = application;
        previousScreenUpdating = application.ScreenUpdating;
        previousEnableEvents = application.EnableEvents;
        previousCalculation = application.Calculation;

        application.ScreenUpdating = false;
        application.EnableEvents = false;
        application.Calculation = ExcelInterop.XlCalculation.xlCalculationManual;
    }

    public void Dispose()
    {
        application.Calculation = previousCalculation;
        application.EnableEvents = previousEnableEvents;
        application.ScreenUpdating = previousScreenUpdating;
    }
}
```

Update `docs/modules/ribbon-sync-current-behavior.md` to document:

- full download now batches data writes by contiguous managed-column segment
- full upload now batch-reads managed regions and only falls back to per-cell text reads for unsafe formatted cells
- partial sync now caches row-ID lookups within one operation

- [ ] **Step 4: Run the full Excel add-in test suite**

Run:

```powershell
dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --no-restore
```

Expected: PASS with `0` failed tests.

- [ ] **Step 5: Run the VSTO add-in build**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File eng\Build-VstoAddIn.ps1 -ProjectPath src\OfficeAgent.ExcelAddIn\OfficeAgent.ExcelAddIn.csproj -Configuration Debug
```

Expected: PASS with `0` errors.

- [ ] **Step 6: Commit the final performance slice**

```powershell
git add src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs docs/modules/ribbon-sync-current-behavior.md
git commit -m "perf: optimize ribbon sync excel io"
```
