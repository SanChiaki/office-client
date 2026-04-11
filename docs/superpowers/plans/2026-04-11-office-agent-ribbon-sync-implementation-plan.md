# OfficeAgent Ribbon Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Ribbon-driven Excel data sync workflow that downloads, partially refreshes, fully uploads, partially uploads, and incrementally uploads workbook data through `/find`, `/head`, and `/batchSave`, without using the task pane or Agent flow.

**Architecture:** Keep Ribbon, native dialogs, and Excel Interop inside `OfficeAgent.ExcelAddIn`, but move sync orchestration, diffing, and preview generation into `OfficeAgent.Core`. Encapsulate the current business system's flattened JSON and mixed-header rules inside a dedicated connector in `OfficeAgent.Infrastructure`, so later systems can plug in behind `systemKey` without rewriting the Ribbon or worksheet logic.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, WinForms, HttpClient, Newtonsoft.Json, xUnit, Node.js mock server

---

## File Structure

- `src/OfficeAgent.Core/Models/ProjectOption.cs`
  Purpose: stable project dropdown item carrying `systemKey`, `projectId`, and display text
- `src/OfficeAgent.Core/Models/WorksheetColumnBinding.cs`
  Purpose: normalized column identity using `apiFieldKey`, header text, activity identity, and ID-column flag
- `src/OfficeAgent.Core/Models/WorksheetSchema.cs`
  Purpose: ordered mixed-header schema for one project sheet
- `src/OfficeAgent.Core/Models/SheetBinding.cs`
  Purpose: workbook sheet to `systemKey/projectId/projectName` binding
- `src/OfficeAgent.Core/Models/WorksheetSnapshotCell.cs`
  Purpose: persisted baseline for `sheetName + rowId + apiFieldKey + value`
- `src/OfficeAgent.Core/Models/SelectedVisibleCell.cs`
  Purpose: visible cell coordinates captured from a possibly non-contiguous selection
- `src/OfficeAgent.Core/Models/ResolvedSelection.cs`
  Purpose: selected `id` values, selected `apiFieldKey` values, and target cells to refresh
- `src/OfficeAgent.Core/Models/CellChange.cs`
  Purpose: one changed cell destined for `/batchSave`
- `src/OfficeAgent.Core/Models/SyncOperationPreview.cs`
  Purpose: dialog-ready summary and diff preview for download and upload actions
- `src/OfficeAgent.Core/Services/ISystemConnector.cs`
  Purpose: unified connector contract for project lookup, schema lookup, find, and batch save
- `src/OfficeAgent.Core/Services/IWorksheetMetadataStore.cs`
  Purpose: save/load visible metadata sheet content
- `src/OfficeAgent.Core/Services/IWorksheetSelectionReader.cs`
  Purpose: read the current workbook selection as visible cell coordinates
- `src/OfficeAgent.Core/Sync/WorksheetChangeTracker.cs`
  Purpose: compare live cells to snapshots and detect overwrite conflicts
- `src/OfficeAgent.Core/Sync/SyncOperationPreviewFactory.cs`
  Purpose: build consistent confirmation payloads for dialogs
- `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
  Purpose: orchestrate the five Ribbon actions against metadata, connector, layout writer, and change tracker
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessHeadDefinition.cs`
  Purpose: deserialize `/head` payloads
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSchemaMapper.cs`
  Purpose: expand flat `/head` + `/find` data into a mixed-header `WorksheetSchema`
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
  Purpose: current-system implementation of `ISystemConnector`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
  Purpose: implement visible `_OfficeAgentMetadata` read/write logic using Excel Interop
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs`
  Purpose: render mixed headers and data into a managed project sheet
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs`
  Purpose: convert visible selected cells into `ResolvedSelection`
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Purpose: bridge Ribbon events to `WorksheetSyncService` and dialogs
- `src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs`
  Purpose: lightweight native confirmation for downloads
- `src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs`
  Purpose: strong native confirmation with diff preview for uploads
- `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
  Purpose: success/error summary dialog
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Purpose: Ribbon event handlers and UI-to-controller delegation
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
  Purpose: project dropdown and five sync buttons
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Purpose: compose sync services and refresh Ribbon state on sheet changes
- `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
  Purpose: explicit compile list for newly added add-in files
- `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSchemaMapperTests.cs`
  Purpose: verify mixed-header schema expansion
- `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`
  Purpose: verify `/find`, `/head`, and `/batchSave` request/response handling
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`
  Purpose: verify visible metadata sheet behavior through a fake adapter
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSchemaLayoutServiceTests.cs`
  Purpose: verify header merge planning
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSelectionResolverTests.cs`
  Purpose: verify selected cells resolve to `id + apiFieldKey` without selecting headers or ID cells
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Purpose: verify Ribbon/controller orchestration
- `tests/OfficeAgent.Core.Tests/WorksheetChangeTrackerTests.cs`
  Purpose: verify dirty-cell detection and overwrite blocking
- `tests/OfficeAgent.Core.Tests/SyncOperationPreviewFactoryTests.cs`
  Purpose: verify dialog summary and diff generation
- `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`
  Purpose: verify the five action paths and snapshot updates
- `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`
  Purpose: verify mock-server roundtrips for find and batch save
- `tests/mock-server/server.js`
  Purpose: provide `/find`, `/head`, and `/batchSave` endpoints for local verification
- `docs/vsto-manual-test-checklist.md`
  Purpose: add manual cases for the new Ribbon sync workflow

### Task 1: Add Sync Contracts and Mixed-Header Schema Mapping

**Files:**
- Create: `src/OfficeAgent.Core/Models/ProjectOption.cs`
- Create: `src/OfficeAgent.Core/Models/WorksheetColumnBinding.cs`
- Create: `src/OfficeAgent.Core/Models/WorksheetSchema.cs`
- Create: `src/OfficeAgent.Core/Models/CellChange.cs`
- Create: `src/OfficeAgent.Core/Services/ISystemConnector.cs`
- Create: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessHeadDefinition.cs`
- Create: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSchemaMapper.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSchemaMapperTests.cs`

- [ ] **Step 1: Write the failing schema-mapping test**

```csharp
using System.Collections.Generic;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class CurrentBusinessSchemaMapperTests
    {
        [Fact]
        public void BuildExpandsActivityFieldsIntoMixedColumns()
        {
            var mapper = new CurrentBusinessSchemaMapper(
                new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                {
                    ["start"] = "开始时间",
                    ["end"] = "结束时间",
                });

            var schema = mapper.Build(
                "performance",
                new[]
                {
                    new CurrentBusinessHeadDefinition { FieldKey = "id", HeaderText = "ID", IsId = true },
                    new CurrentBusinessHeadDefinition { FieldKey = "name", HeaderText = "项目名称" },
                    new CurrentBusinessHeadDefinition { HeadType = "activity", ActivityId = "12345678", ActivityName = "测试活动111" },
                },
                new[]
                {
                    new Dictionary<string, object>
                    {
                        ["id"] = "row-1",
                        ["name"] = "项目A",
                        ["start_12345678"] = "2026-01-02",
                        ["end_12345678"] = "2026-01-03",
                    },
                });

            Assert.Collection(
                schema.Columns,
                column => Assert.Equal("id", column.ApiFieldKey),
                column => Assert.Equal("name", column.ApiFieldKey),
                column =>
                {
                    Assert.Equal("start_12345678", column.ApiFieldKey);
                    Assert.Equal(WorksheetColumnKind.ActivityProperty, column.ColumnKind);
                    Assert.Equal("测试活动111", column.ParentHeaderText);
                    Assert.Equal("开始时间", column.ChildHeaderText);
                },
                column =>
                {
                    Assert.Equal("end_12345678", column.ApiFieldKey);
                    Assert.Equal(WorksheetColumnKind.ActivityProperty, column.ColumnKind);
                    Assert.Equal("测试活动111", column.ParentHeaderText);
                    Assert.Equal("结束时间", column.ChildHeaderText);
                });
        }
    }
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSchemaMapperTests.BuildExpandsActivityFieldsIntoMixedColumns`

Expected: FAIL with compiler errors mentioning `CurrentBusinessSchemaMapper`, `CurrentBusinessHeadDefinition`, or `WorksheetColumnKind` missing.

- [ ] **Step 3: Write the minimal schema models and mapper**

```csharp
namespace OfficeAgent.Core.Models
{
    public enum WorksheetColumnKind
    {
        Single,
        ActivityProperty,
    }

    public sealed class WorksheetColumnBinding
    {
        public int ColumnIndex { get; set; }
        public string ApiFieldKey { get; set; } = string.Empty;
        public WorksheetColumnKind ColumnKind { get; set; }
        public string ParentHeaderText { get; set; } = string.Empty;
        public string ChildHeaderText { get; set; } = string.Empty;
        public string ActivityId { get; set; } = string.Empty;
        public string ActivityName { get; set; } = string.Empty;
        public string PropertyKey { get; set; } = string.Empty;
        public bool IsIdColumn { get; set; }
    }

    public sealed class WorksheetSchema
    {
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public WorksheetColumnBinding[] Columns { get; set; } = System.Array.Empty<WorksheetColumnBinding>();
    }
}
```

```csharp
namespace OfficeAgent.Infrastructure.Http
{
    public sealed class CurrentBusinessHeadDefinition
    {
        public string FieldKey { get; set; } = string.Empty;
        public string HeaderText { get; set; } = string.Empty;
        public string HeadType { get; set; } = string.Empty;
        public string ActivityId { get; set; } = string.Empty;
        public string ActivityName { get; set; } = string.Empty;
        public bool IsId { get; set; }
    }
}
```

```csharp
public sealed class CurrentBusinessSchemaMapper
{
    private readonly IReadOnlyDictionary<string, string> propertyNames;

    public CurrentBusinessSchemaMapper(IReadOnlyDictionary<string, string> propertyNames)
    {
        this.propertyNames = propertyNames;
    }

    public WorksheetSchema Build(
        string projectId,
        IReadOnlyList<CurrentBusinessHeadDefinition> headList,
        IReadOnlyList<IDictionary<string, object>> rows)
    {
        var columns = new List<WorksheetColumnBinding>();
        var nextColumn = 1;

        foreach (var head in headList.Where((item) => !string.Equals(item.HeadType, "activity", StringComparison.OrdinalIgnoreCase)))
        {
            columns.Add(new WorksheetColumnBinding
            {
                ColumnIndex = nextColumn++,
                ApiFieldKey = head.FieldKey,
                ColumnKind = WorksheetColumnKind.Single,
                ParentHeaderText = head.HeaderText,
                ChildHeaderText = head.HeaderText,
                IsIdColumn = head.IsId,
            });
        }

        var activityHeads = headList
            .Where((item) => string.Equals(item.HeadType, "activity", StringComparison.OrdinalIgnoreCase))
            .ToDictionary((item) => item.ActivityId, StringComparer.OrdinalIgnoreCase);

        var flatKeys = rows.SelectMany((row) => row.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
        foreach (var flatKey in flatKeys.Where((key) => key.Contains("_")))
        {
            var segments = flatKey.Split(new[] { '_' }, 2);
            if (segments.Length != 2 || !activityHeads.TryGetValue(segments[1], out var activityHead))
            {
                continue;
            }

            var propertyKey = segments[0];
            columns.Add(new WorksheetColumnBinding
            {
                ColumnIndex = nextColumn++,
                ApiFieldKey = flatKey,
                ColumnKind = WorksheetColumnKind.ActivityProperty,
                ParentHeaderText = activityHead.ActivityName,
                ChildHeaderText = propertyNames.TryGetValue(propertyKey, out var label) ? label : propertyKey,
                ActivityId = activityHead.ActivityId,
                ActivityName = activityHead.ActivityName,
                PropertyKey = propertyKey,
            });
        }

        return new WorksheetSchema
        {
            SystemKey = "current-business-system",
            ProjectId = projectId,
            Columns = columns.ToArray(),
        };
    }
}
```

- [ ] **Step 4: Run the targeted test and the infrastructure suite**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSchemaMapperTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`

Expected: PASS with the new schema-mapping test included.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSchemaMapperTests.cs src/OfficeAgent.Core/Models/ProjectOption.cs src/OfficeAgent.Core/Models/WorksheetColumnBinding.cs src/OfficeAgent.Core/Models/WorksheetSchema.cs src/OfficeAgent.Core/Models/CellChange.cs src/OfficeAgent.Core/Services/ISystemConnector.cs src/OfficeAgent.Infrastructure/Http/CurrentBusinessHeadDefinition.cs src/OfficeAgent.Infrastructure/Http/CurrentBusinessSchemaMapper.cs
git commit -m "feat: add ribbon sync schema contracts"
```

### Task 2: Add Visible Metadata Worksheet Storage

**Files:**
- Create: `src/OfficeAgent.Core/Models/SheetBinding.cs`
- Create: `src/OfficeAgent.Core/Models/WorksheetSnapshotCell.cs`
- Create: `src/OfficeAgent.Core/Services/IWorksheetMetadataStore.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetMetadataAdapter.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`

- [ ] **Step 1: Write the failing metadata-store test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Excel;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetMetadataStoreTests
    {
        [Fact]
        public void SaveBindingCreatesVisibleMetadataWorksheetAndRoundTripsBinding()
        {
            var adapter = new FakeWorksheetMetadataAdapter();
            var store = new WorksheetMetadataStore(adapter);
            var binding = new SheetBinding
            {
                SheetName = "Sync-performance",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };

            store.SaveBinding(binding);

            Assert.Equal("_OfficeAgentMetadata", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = store.LoadBinding("Sync-performance");
            Assert.Equal("performance", loaded.ProjectId);
            Assert.Equal("绩效项目", loaded.ProjectName);
        }
    }
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetMetadataStoreTests.SaveBindingCreatesVisibleMetadataWorksheetAndRoundTripsBinding`

Expected: FAIL with compiler errors mentioning `WorksheetMetadataStore`, `SheetBinding`, or `IWorksheetMetadataAdapter` missing.

- [ ] **Step 3: Write the minimal metadata models and store**

```csharp
namespace OfficeAgent.Core.Models
{
    public sealed class SheetBinding
    {
        public string SheetName { get; set; } = string.Empty;
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
    }

    public sealed class WorksheetSnapshotCell
    {
        public string SheetName { get; set; } = string.Empty;
        public string RowId { get; set; } = string.Empty;
        public string ApiFieldKey { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }
}
```

```csharp
public interface IWorksheetMetadataStore
{
    void SaveBinding(SheetBinding binding);
    SheetBinding LoadBinding(string sheetName);
    WorksheetSnapshotCell[] LoadSnapshot(string sheetName);
    void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells);
}
```

```csharp
internal interface IWorksheetMetadataAdapter
{
    void EnsureWorksheet(string name, bool visible);
    void WriteTable(string tableName, string[] headers, string[][] rows);
    string[][] ReadTable(string tableName);
}
```

```csharp
internal sealed class WorksheetMetadataStore : IWorksheetMetadataStore
{
    private readonly IWorksheetMetadataAdapter adapter;

    public WorksheetMetadataStore(IWorksheetMetadataAdapter adapter)
    {
        this.adapter = adapter;
    }

    public void SaveBinding(SheetBinding binding)
    {
        adapter.EnsureWorksheet("_OfficeAgentMetadata", visible: true);
        adapter.WriteTable(
            "SheetBindings",
            new[] { "SheetName", "SystemKey", "ProjectId", "ProjectName" },
            new[]
            {
                new[]
                {
                    binding.SheetName,
                    binding.SystemKey,
                    binding.ProjectId,
                    binding.ProjectName,
                },
            });
    }

    public SheetBinding LoadBinding(string sheetName)
    {
        var rows = adapter.ReadTable("SheetBindings");
        var row = rows.First((item) => string.Equals(item[0], sheetName, StringComparison.OrdinalIgnoreCase));
        return new SheetBinding
        {
            SheetName = row[0],
            SystemKey = row[1],
            ProjectId = row[2],
            ProjectName = row[3],
        };
    }
}
```

```xml
<Compile Include="Excel\IWorksheetMetadataAdapter.cs" />
<Compile Include="Excel\ExcelWorkbookMetadataAdapter.cs" />
<Compile Include="Excel\WorksheetMetadataStore.cs" />
```

- [ ] **Step 4: Run the targeted test and the add-in test suite**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetMetadataStoreTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS with the new metadata-store coverage included.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs src/OfficeAgent.Core/Models/SheetBinding.cs src/OfficeAgent.Core/Models/WorksheetSnapshotCell.cs src/OfficeAgent.Core/Services/IWorksheetMetadataStore.cs src/OfficeAgent.ExcelAddIn/Excel/IWorksheetMetadataAdapter.cs src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj
git commit -m "feat: add visible worksheet metadata store"
```

### Task 3: Add Mixed-Header Layout Planning and Writing

**Files:**
- Create: `src/OfficeAgent.Core/Models/HeaderCellPlan.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSchemaLayoutServiceTests.cs`

- [ ] **Step 1: Write the failing layout-planning test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Excel;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetSchemaLayoutServiceTests
    {
        [Fact]
        public void BuildHeaderPlanMergesSingleColumnsVerticallyAndActivityColumnsHorizontally()
        {
            var service = new WorksheetSchemaLayoutService();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", ColumnKind = WorksheetColumnKind.Single, ParentHeaderText = "ID", ChildHeaderText = "ID", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name", ColumnKind = WorksheetColumnKind.Single, ParentHeaderText = "项目名称", ChildHeaderText = "项目名称" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_12345678", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "测试活动111", ChildHeaderText = "开始时间" },
                    new WorksheetColumnBinding { ColumnIndex = 4, ApiFieldKey = "end_12345678", ColumnKind = WorksheetColumnKind.ActivityProperty, ParentHeaderText = "测试活动111", ChildHeaderText = "结束时间" },
                },
            };

            var plan = service.BuildHeaderPlan(schema);

            Assert.Contains(plan, cell => cell.Row == 1 && cell.Column == 1 && cell.RowSpan == 2 && cell.Text == "ID");
            Assert.Contains(plan, cell => cell.Row == 1 && cell.Column == 3 && cell.ColumnSpan == 2 && cell.Text == "测试活动111");
            Assert.Contains(plan, cell => cell.Row == 2 && cell.Column == 4 && cell.Text == "结束时间");
        }
    }
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSchemaLayoutServiceTests.BuildHeaderPlanMergesSingleColumnsVerticallyAndActivityColumnsHorizontally`

Expected: FAIL with compiler errors mentioning `HeaderCellPlan` or `BuildHeaderPlan` missing.

- [ ] **Step 3: Write the minimal header-plan model and layout service**

```csharp
namespace OfficeAgent.Core.Models
{
    public sealed class HeaderCellPlan
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColumnSpan { get; set; } = 1;
        public string Text { get; set; } = string.Empty;
    }
}
```

```csharp
internal sealed class WorksheetSchemaLayoutService
{
    public HeaderCellPlan[] BuildHeaderPlan(WorksheetSchema schema)
    {
        var cells = new List<HeaderCellPlan>();

        foreach (var column in schema.Columns.Where((item) => item.ColumnKind == WorksheetColumnKind.Single))
        {
            cells.Add(new HeaderCellPlan
            {
                Row = 1,
                Column = column.ColumnIndex,
                RowSpan = 2,
                Text = column.ChildHeaderText,
            });
        }

        foreach (var group in schema.Columns.Where((item) => item.ColumnKind == WorksheetColumnKind.ActivityProperty).GroupBy((item) => item.ParentHeaderText))
        {
            var ordered = group.OrderBy((item) => item.ColumnIndex).ToArray();
            cells.Add(new HeaderCellPlan
            {
                Row = 1,
                Column = ordered[0].ColumnIndex,
                ColumnSpan = ordered.Length,
                Text = group.Key,
            });

            foreach (var column in ordered)
            {
                cells.Add(new HeaderCellPlan
                {
                    Row = 2,
                    Column = column.ColumnIndex,
                    Text = column.ChildHeaderText,
                });
            }
        }

        return cells.OrderBy((item) => item.Row).ThenBy((item) => item.Column).ToArray();
    }
}
```

```xml
<Compile Include="Excel\WorksheetSchemaLayoutService.cs" />
```

- [ ] **Step 4: Run the targeted test and the add-in tests**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSchemaLayoutServiceTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS with the new layout-planning coverage included.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSchemaLayoutServiceTests.cs src/OfficeAgent.Core/Models/HeaderCellPlan.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj
git commit -m "feat: add mixed header layout planning"
```

### Task 4: Add Visible Selection Resolution Without Requiring Header or ID Selection

**Files:**
- Create: `src/OfficeAgent.Core/Models/SelectedVisibleCell.cs`
- Create: `src/OfficeAgent.Core/Models/ResolvedSelection.cs`
- Create: `src/OfficeAgent.Core/Services/IWorksheetSelectionReader.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSelectionResolverTests.cs`

- [ ] **Step 1: Write the failing selection-resolution test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Excel;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetSelectionResolverTests
    {
        [Fact]
        public void ResolveReturnsIdsAndFieldKeysWhenHeadersAndIdCellsAreNotSelected()
        {
            var resolver = new WorksheetSelectionResolver();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_12345678" },
                },
            };

            var resolved = resolver.Resolve(
                schema,
                new[]
                {
                    new SelectedVisibleCell { Row = 3, Column = 2, Value = "项目A" },
                    new SelectedVisibleCell { Row = 3, Column = 3, Value = "2026-01-02" },
                },
                row => row == 3 ? "row-1" : string.Empty);

            Assert.Equal(new[] { "row-1" }, resolved.RowIds);
            Assert.Equal(new[] { "name", "start_12345678" }, resolved.ApiFieldKeys);
            Assert.Equal(2, resolved.TargetCells.Length);
        }
    }
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSelectionResolverTests.ResolveReturnsIdsAndFieldKeysWhenHeadersAndIdCellsAreNotSelected`

Expected: FAIL with compiler errors mentioning `SelectedVisibleCell`, `ResolvedSelection`, or `WorksheetSelectionResolver` missing.

- [ ] **Step 3: Write the minimal selection models and resolver**

```csharp
namespace OfficeAgent.Core.Models
{
    public sealed class SelectedVisibleCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Value { get; set; } = string.Empty;
    }

    public sealed class ResolvedSelection
    {
        public string[] RowIds { get; set; } = System.Array.Empty<string>();
        public string[] ApiFieldKeys { get; set; } = System.Array.Empty<string>();
        public SelectedVisibleCell[] TargetCells { get; set; } = System.Array.Empty<SelectedVisibleCell>();
    }
}
```

```csharp
internal sealed class WorksheetSelectionResolver
{
    public ResolvedSelection Resolve(
        WorksheetSchema schema,
        IReadOnlyList<SelectedVisibleCell> visibleCells,
        System.Func<int, string> rowIdAccessor)
    {
        var rowIds = visibleCells
            .Select((cell) => rowIdAccessor(cell.Row))
            .Where((value) => !string.IsNullOrWhiteSpace(value))
            .Distinct(System.StringComparer.Ordinal)
            .ToArray();

        var fieldKeys = visibleCells
            .Select((cell) => schema.Columns.First((column) => column.ColumnIndex == cell.Column).ApiFieldKey)
            .Distinct(System.StringComparer.Ordinal)
            .ToArray();

        return new ResolvedSelection
        {
            RowIds = rowIds,
            ApiFieldKeys = fieldKeys,
            TargetCells = visibleCells.ToArray(),
        };
    }
}
```

```xml
<Compile Include="Excel\WorksheetSelectionResolver.cs" />
```

- [ ] **Step 4: Run the targeted test and the add-in tests**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSelectionResolverTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS with the new resolver coverage included.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSelectionResolverTests.cs src/OfficeAgent.Core/Models/SelectedVisibleCell.cs src/OfficeAgent.Core/Models/ResolvedSelection.cs src/OfficeAgent.Core/Services/IWorksheetSelectionReader.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj
git commit -m "feat: resolve visible selections for ribbon sync"
```

### Task 5: Add Dirty-Cell Tracking and Confirmation Preview Generation

**Files:**
- Create: `src/OfficeAgent.Core/Models/SyncOperationPreview.cs`
- Create: `src/OfficeAgent.Core/Sync/WorksheetChangeTracker.cs`
- Create: `src/OfficeAgent.Core/Sync/SyncOperationPreviewFactory.cs`
- Test: `tests/OfficeAgent.Core.Tests/WorksheetChangeTrackerTests.cs`
- Test: `tests/OfficeAgent.Core.Tests/SyncOperationPreviewFactoryTests.cs`

- [ ] **Step 1: Write the failing dirty-cell test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class WorksheetChangeTrackerTests
    {
        [Fact]
        public void GetDirtyCellsReturnsOnlyChangedCellsForExistingIds()
        {
            var tracker = new WorksheetChangeTracker();
            var dirty = tracker.GetDirtyCells(
                "Sync-performance",
                new[]
                {
                    new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "旧值" },
                    new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "start_12345678", Value = "2026-01-01" },
                },
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "旧值", NewValue = "新值" },
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "start_12345678", OldValue = "2026-01-01", NewValue = "2026-01-01" },
                });

            var changed = Assert.Single(dirty);
            Assert.Equal("name", changed.ApiFieldKey);
            Assert.Equal("新值", changed.NewValue);
        }
    }
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~WorksheetChangeTrackerTests.GetDirtyCellsReturnsOnlyChangedCellsForExistingIds`

Expected: FAIL with compiler errors mentioning `WorksheetChangeTracker` or `SyncOperationPreview` missing.

- [ ] **Step 3: Write the minimal tracker and preview factory**

```csharp
namespace OfficeAgent.Core.Models
{
    public sealed class SyncOperationPreview
    {
        public string OperationName { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
        public string[] Details { get; set; } = System.Array.Empty<string>();
        public CellChange[] Changes { get; set; } = System.Array.Empty<CellChange>();
    }
}
```

```csharp
public sealed class WorksheetChangeTracker
{
    public CellChange[] GetDirtyCells(
        string sheetName,
        IReadOnlyList<WorksheetSnapshotCell> snapshot,
        IReadOnlyList<CellChange> currentCells)
    {
        var baseline = snapshot.ToDictionary(
            item => $"{item.RowId}|{item.ApiFieldKey}",
            item => item.Value,
            StringComparer.Ordinal);

        return currentCells
            .Where(item => string.Equals(item.SheetName, sheetName, StringComparison.Ordinal))
            .Where(item => baseline.TryGetValue($"{item.RowId}|{item.ApiFieldKey}", out var oldValue) &&
                           !string.Equals(oldValue, item.NewValue, StringComparison.Ordinal))
            .ToArray();
    }
}
```

```csharp
public sealed class SyncOperationPreviewFactory
{
    public SyncOperationPreview CreateUploadPreview(string operationName, IReadOnlyList<CellChange> changes)
    {
        return new SyncOperationPreview
        {
            OperationName = operationName,
            Summary = $"Upload {changes.Count} changed cell(s).",
            Details = changes.Take(3).Select(item => $"{item.RowId} / {item.ApiFieldKey}: {item.OldValue} -> {item.NewValue}").ToArray(),
            Changes = changes.ToArray(),
        };
    }
}
```

- [ ] **Step 4: Run the targeted tests and the core suite**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~WorksheetChangeTrackerTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~SyncOperationPreviewFactoryTests`

Expected: PASS after adding the companion preview-factory test.

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`

Expected: PASS with dirty tracking and preview coverage included.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.Core.Tests/WorksheetChangeTrackerTests.cs tests/OfficeAgent.Core.Tests/SyncOperationPreviewFactoryTests.cs src/OfficeAgent.Core/Models/SyncOperationPreview.cs src/OfficeAgent.Core/Sync/WorksheetChangeTracker.cs src/OfficeAgent.Core/Sync/SyncOperationPreviewFactory.cs
git commit -m "feat: add dirty tracking for ribbon sync"
```

### Task 6: Add `/find`, `/head`, and `/batchSave` Connector Support and Mock Server Endpoints

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessBatchSaveItem.cs`
- Create: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- Modify: `tests/mock-server/server.js`
- Test: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`
- Test: `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`

- [ ] **Step 1: Write the failing connector-request test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class CurrentBusinessSystemConnectorTests
    {
        [Fact]
        public void BatchSaveSendsOneItemPerChangedCell()
        {
            var handler = new RecordingHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            connector.BatchSave(
                "performance",
                new[]
                {
                    new CellChange
                    {
                        RowId = "row-1",
                        ApiFieldKey = "start_12345678",
                        NewValue = "2026-01-02",
                    },
                });

            Assert.Equal("https://api.internal.example/batchSave", handler.LastRequestUri);
            Assert.Contains("\"fieldKey\":\"start_12345678\"", handler.LastBody);
            Assert.Contains("\"value\":\"2026-01-02\"", handler.LastBody);
        }
    }
}
```

- [ ] **Step 2: Run the connector test to verify it fails**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorTests.BatchSaveSendsOneItemPerChangedCell`

Expected: FAIL with compiler errors mentioning `CurrentBusinessSystemConnector` or `BatchSave` missing.

- [ ] **Step 3: Write the minimal connector and mock endpoints**

```csharp
public sealed class CurrentBusinessBatchSaveItem
{
    public string ProjectId { get; set; } = string.Empty;
    public string Id { get; set; } = string.Empty;
    public string FieldKey { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
```

```csharp
public sealed class CurrentBusinessSystemConnector : ISystemConnector
{
    public IReadOnlyList<ProjectOption> GetProjects()
    {
        return new[]
        {
            new ProjectOption { SystemKey = "current-business-system", ProjectId = "performance", DisplayName = "绩效项目" },
        };
    }

    public WorksheetSchema GetSchema(string projectId)
    {
        var headList = GetHeadDefinitions(projectId);
        var rows = Find(projectId, System.Array.Empty<string>(), System.Array.Empty<string>());
        return schemaMapper.Build(projectId, headList, rows);
    }

    public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
    {
        var endpoint = new Uri($"{AppSettings.NormalizeBaseUrl(loadSettings().BaseUrl)}/find");
        var payload = JsonConvert.SerializeObject(new { projectId = projectId, ids = rowIds, fieldKeys = fieldKeys });
        using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
        {
            request.Content = new StringContent(payload, Encoding.UTF8, "application/json");
            using (var response = httpClient.SendAsync(request).GetAwaiter().GetResult())
            {
                response.EnsureSuccessStatusCode();
                var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(body);
            }
        }
    }

    public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
    {
        var payload = changes.Select(item => new CurrentBusinessBatchSaveItem
        {
            ProjectId = projectId,
            Id = item.RowId,
            FieldKey = item.ApiFieldKey,
            Value = item.NewValue,
        }).ToArray();

        var endpoint = new Uri($"{AppSettings.NormalizeBaseUrl(loadSettings().BaseUrl)}/batchSave");
        using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
        {
            request.Content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
            using (var response = httpClient.SendAsync(request).GetAwaiter().GetResult())
            {
                response.EnsureSuccessStatusCode();
            }
        }
    }
}
```

```javascript
apiApp.post('/head', requireAuth, function (_req, res) {
  return res.json({
    headList: [
      { fieldKey: 'id', headerText: 'ID', isId: true },
      { fieldKey: 'name', headerText: '项目名称' },
      { headType: 'activity', activityId: '12345678', activityName: '测试活动111' },
    ],
  });
});

apiApp.post('/find', requireAuth, function (req, res) {
  var ids = ((req.body || {}).ids) || [];
  var fieldKeys = ((req.body || {}).fieldKeys) || [];
  var rows = performanceRows.filter(function (row) {
    return ids.length === 0 || ids.indexOf(row.id) >= 0;
  }).map(function (row) {
    if (fieldKeys.length === 0) return row;
    var projected = { id: row.id };
    fieldKeys.forEach(function (key) { projected[key] = row[key]; });
    return projected;
  });
  res.json(rows);
});

apiApp.post('/batchSave', requireAuth, function (req, res) {
  var changes = Array.isArray(req.body) ? req.body : [];
  changes.forEach(function (item) {
    var target = performanceRows.find(function (row) { return row.id === item.id; });
    if (target) {
      target[item.fieldKey] = item.value;
    }
  });
  res.json({ savedCount: changes.length, message: '批量保存成功。' });
});
```

- [ ] **Step 4: Run the targeted tests and the integration suite**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorIntegrationTests`

Expected: PASS with mock-server-backed `/find` and `/batchSave` coverage.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs tests/mock-server/server.js src/OfficeAgent.Infrastructure/Http/CurrentBusinessBatchSaveItem.cs src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs
git commit -m "feat: add current business ribbon sync connector"
```

### Task 7: Add Worksheet Sync Orchestration for All Five Ribbon Actions

**Files:**
- Create: `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- Test: `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`

- [ ] **Step 1: Write the failing orchestration test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class WorksheetSyncServiceTests
    {
        [Fact]
        public void PrepareIncrementalUploadReturnsOnlyDirtyCellsForExistingRows()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var tracker = new WorksheetChangeTracker();
            var previewFactory = new SyncOperationPreviewFactory();
            var service = new WorksheetSyncService(connector, metadataStore, tracker, previewFactory);

            metadataStore.SaveSnapshot("Sync-performance", new[]
            {
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "旧值" },
            });

            var preview = service.PrepareIncrementalUpload(
                "Sync-performance",
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "旧值", NewValue = "新值" },
                    new CellChange { SheetName = "Sync-performance", RowId = string.Empty, ApiFieldKey = "name", OldValue = string.Empty, NewValue = "忽略" },
                });

            var change = Assert.Single(preview.Changes);
            Assert.Equal("row-1", change.RowId);
            Assert.Equal("name", change.ApiFieldKey);
        }
    }
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~WorksheetSyncServiceTests.PrepareIncrementalUploadReturnsOnlyDirtyCellsForExistingRows`

Expected: FAIL with compiler errors mentioning `WorksheetSyncService` missing.

- [ ] **Step 3: Write the minimal orchestration service**

```csharp
public sealed class WorksheetSyncService
{
    private readonly ISystemConnector connector;
    private readonly IWorksheetMetadataStore metadataStore;
    private readonly WorksheetChangeTracker changeTracker;
    private readonly SyncOperationPreviewFactory previewFactory;

    public WorksheetSyncService(
        ISystemConnector connector,
        IWorksheetMetadataStore metadataStore,
        WorksheetChangeTracker changeTracker,
        SyncOperationPreviewFactory previewFactory)
    {
        this.connector = connector;
        this.metadataStore = metadataStore;
        this.changeTracker = changeTracker;
        this.previewFactory = previewFactory;
    }

    public SyncOperationPreview PrepareIncrementalUpload(string sheetName, IReadOnlyList<CellChange> currentCells)
    {
        var snapshot = metadataStore.LoadSnapshot(sheetName);
        var dirtyCells = changeTracker
            .GetDirtyCells(sheetName, snapshot, currentCells)
            .Where(item => !string.IsNullOrWhiteSpace(item.RowId))
            .ToArray();

        return previewFactory.CreateUploadPreview("增量上传", dirtyCells);
    }

    public WorksheetSchema LoadSchemaForSheet(string sheetName)
    {
        var binding = metadataStore.LoadBinding(sheetName);
        return connector.GetSchema(binding.ProjectId);
    }

    public IReadOnlyList<IDictionary<string, object>> ExecutePartialDownload(string sheetName, ResolvedSelection selection)
    {
        var binding = metadataStore.LoadBinding(sheetName);
        return connector.Find(binding.ProjectId, selection.RowIds, selection.ApiFieldKeys);
    }

    public void ExecutePartialUpload(string sheetName, IReadOnlyList<CellChange> changes)
    {
        var binding = metadataStore.LoadBinding(sheetName);
        connector.BatchSave(binding.ProjectId, changes);
    }
}
```

- [ ] **Step 4: Run the targeted test and the core suite**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~WorksheetSyncServiceTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`

Expected: PASS with the orchestration test included.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs src/OfficeAgent.Core/Sync/WorksheetSyncService.cs
git commit -m "feat: add ribbon sync orchestration service"
```

### Task 8: Add Native Dialogs and Ribbon Controller Integration

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`

- [ ] **Step 1: Write the failing controller test**

```csharp
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class RibbonSyncControllerTests
    {
        [Fact]
        public void SelectProjectSavesBindingAndRefreshesRibbonState()
        {
            var syncService = new FakeWorksheetSyncService();
            var metadataStore = new FakeWorksheetMetadataStore();
            var controller = new RibbonSyncController(syncService, metadataStore);

            controller.SelectProject(
                "Sync-performance",
                new ProjectOption
                {
                    SystemKey = "current-business-system",
                    ProjectId = "performance",
                    DisplayName = "绩效项目",
                });

            var binding = metadataStore.LoadBinding("Sync-performance");
            Assert.Equal("performance", binding.ProjectId);
            Assert.True(controller.HasActiveProject);
            Assert.Equal("绩效项目", controller.CurrentProjectDisplayName);
        }
    }
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests.SelectProjectSavesBindingAndRefreshesRibbonState`

Expected: FAIL with compiler errors mentioning `RibbonSyncController` missing.

- [ ] **Step 3: Write the minimal controller, dialogs, and Ribbon wiring**

```csharp
internal sealed class RibbonSyncController
{
    private readonly WorksheetSyncService syncService;
    private readonly IWorksheetMetadataStore metadataStore;

    public RibbonSyncController(WorksheetSyncService syncService, IWorksheetMetadataStore metadataStore)
    {
        this.syncService = syncService;
        this.metadataStore = metadataStore;
    }

    public bool HasActiveProject { get; private set; }
    public string CurrentProjectDisplayName { get; private set; } = "先选择项目";

    public void SelectProject(string sheetName, ProjectOption project)
    {
        metadataStore.SaveBinding(new SheetBinding
        {
            SheetName = sheetName,
            SystemKey = project.SystemKey,
            ProjectId = project.ProjectId,
            ProjectName = project.DisplayName,
        });

        HasActiveProject = true;
        CurrentProjectDisplayName = project.DisplayName;
    }
}
```

```csharp
private RibbonSyncController ribbonSyncController;

private void ThisAddIn_Startup(object sender, EventArgs e)
{
    var worksheetMetadataStore = new WorksheetMetadataStore(new ExcelWorkbookMetadataAdapter(Application));
    var worksheetSyncService = new WorksheetSyncService(
        currentBusinessSystemConnector,
        worksheetMetadataStore,
        new WorksheetChangeTracker(),
        new SyncOperationPreviewFactory());
    ribbonSyncController = new RibbonSyncController(worksheetSyncService, worksheetMetadataStore);
}
```

```csharp
this.projectGroup = Factory.CreateRibbonGroup();
this.projectDropDown = Factory.CreateRibbonDropDown();
this.downloadGroup = Factory.CreateRibbonGroup();
this.fullDownloadButton = Factory.CreateRibbonButton();
this.partialDownloadButton = Factory.CreateRibbonButton();
this.uploadGroup = Factory.CreateRibbonGroup();
this.fullUploadButton = Factory.CreateRibbonButton();
this.partialUploadButton = Factory.CreateRibbonButton();
this.incrementalUploadButton = Factory.CreateRibbonButton();
```

```xml
<Compile Include="RibbonSyncController.cs" />
<Compile Include="Dialogs\DownloadConfirmDialog.cs" />
<Compile Include="Dialogs\UploadConfirmDialog.cs" />
<Compile Include="Dialogs\OperationResultDialog.cs" />
```

- [ ] **Step 4: Run the targeted test, the add-in tests, and build the add-in**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS with the controller coverage included.

Run: `& "$env:ProgramFiles\\Microsoft Visual Studio\\2022\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe" "src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj" /restore /p:RestorePackagesConfig=true /p:Configuration=Debug`

Expected: BUILD SUCCEEDED.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj
git commit -m "feat: wire ribbon sync actions and dialogs"
```

### Task 9: Add Roundtrip Integration Coverage and Manual Verification Notes

**Files:**
- Create: `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Write the failing roundtrip integration test**

```csharp
using System.Linq;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.IntegrationTests
{
    public sealed class CurrentBusinessSystemConnectorIntegrationTests : IClassFixture<MockServerFixture>
    {
        private readonly MockServerFixture fixture;

        public CurrentBusinessSystemConnectorIntegrationTests(MockServerFixture fixture)
        {
            this.fixture = fixture;
        }

        [Fact]
        public async Task FindAndBatchSaveRoundTripAgainstMockServer()
        {
            var connector = await fixture.CreateCurrentBusinessConnector();

            var rows = connector.Find("performance", new string[0], new[] { "name", "start_12345678" });
            Assert.NotEmpty(rows);

            connector.BatchSave(
                "performance",
                new[]
                {
                    new CellChange { RowId = rows.First()["id"].ToString(), ApiFieldKey = "name", NewValue = "已修改项目" },
                });

            var afterSave = connector.Find("performance", new string[0], new[] { "name" });
            Assert.Contains(afterSave, row => row["name"].ToString() == "已修改项目");
        }
    }
}
```

- [ ] **Step 2: Run the integration test to verify it fails**

Run: `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorIntegrationTests.FindAndBatchSaveRoundTripAgainstMockServer`

Expected: FAIL until the connector helper on `MockServerFixture` and the new endpoints exist.

- [ ] **Step 3: Add the integration helper and update the manual checklist**

```csharp
public async Task<CurrentBusinessSystemConnector> CreateCurrentBusinessConnector()
{
    var cookieJar = await LoginAs("sync_user", "password123");
    return new CurrentBusinessSystemConnector(
        () => new AppSettings { BaseUrl = BusinessUrl, ApiKey = string.Empty },
        cookieJar);
}
```

```markdown
## Ribbon Sync

- Use the Ribbon project dropdown to bind a project to a blank sheet and confirm the dropdown rehydrates when you switch away and back.
- Run `全量下载` into a managed project sheet and confirm `_OfficeAgentMetadata` stays visible for debugging.
- Select a non-contiguous visible range that excludes both the header rows and the ID column, then run `部分下载`.
- Edit one existing ID row and confirm `增量上传` only previews the dirty cell changes.
- Edit a dirty cell and then run `全量下载`; confirm the warning dialog defaults to cancel.
- Run `部分上传` on one activity-property cell and confirm the payload updates only that `fieldKey`.
```

- [ ] **Step 4: Run the integration suite and the documented verification commands**

Run: `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs docs/vsto-manual-test-checklist.md
git commit -m "test: add ribbon sync roundtrip coverage"
```

## Self-Review

- Spec coverage:
  - Ribbon dropdown and five buttons: Task 8
  - Visible metadata sheet: Task 2
  - Mixed single + activity headers: Tasks 1 and 3
  - Visible/non-contiguous partial selection: Task 4
  - Dirty-cell overwrite blocking and incremental upload: Tasks 5 and 7
  - `/find`, `/head`, `/batchSave`: Task 6
  - Mock server and manual checklist: Task 9
- Placeholder scan:
  - No placeholder markers or vague “same as above” instructions remain
- Type consistency:
  - `WorksheetColumnKind`, `WorksheetColumnBinding`, `ResolvedSelection`, `CellChange`, `SyncOperationPreview`, and `WorksheetSyncService` names are used consistently across tasks
