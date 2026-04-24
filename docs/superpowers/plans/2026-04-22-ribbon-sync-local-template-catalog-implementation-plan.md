# Ribbon Sync Local Template Catalog Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a local per-project template catalog for Ribbon Sync so users can apply, save back, and save-as multiple `AI_Setting` templates while keeping `AI_Setting` itself editable as the live worksheet copy.

**Architecture:** Keep `AI_Setting` as the runtime source of truth for download/upload/initialize, but add a local JSON-backed template asset store plus a new `TemplateBindings` metadata section that records which template the current sheet is derived from. Implement template orchestration in `OfficeAgent.Core`, persist template assets in `OfficeAgent.Infrastructure`, and keep Excel/VSTO concerns in a dedicated `RibbonTemplateController` plus small WinForms dialogs.

**Tech Stack:** C#, .NET Framework 4.8, VSTO Ribbon, WinForms, Newtonsoft.Json, xUnit

---

## File Structure

- `src/OfficeAgent.Core/Models/TemplateDefinition.cs`
  Responsibility: represent one persisted template asset, including layout parameters, mapping-definition snapshot, mapping rows, revision, and timestamps.
- `src/OfficeAgent.Core/Models/TemplateFieldMappingRow.cs`
  Responsibility: hold one template-scoped mapping row without persisting worksheet-specific `SheetName`.
- `src/OfficeAgent.Core/Models/SheetTemplateBinding.cs`
  Responsibility: represent one worksheet row in `AI_Setting.TemplateBindings`.
- `src/OfficeAgent.Core/Models/SheetTemplateState.cs`
  Responsibility: expose the controller-facing template state for the active sheet, including button enablement and dirty state.
- `src/OfficeAgent.Core/Services/IWorksheetTemplateBindingStore.cs`
  Responsibility: abstract read/write of worksheet-scoped template metadata without widening the existing `IWorksheetMetadataStore` contract.
- `src/OfficeAgent.Core/Services/ITemplateStore.cs`
  Responsibility: abstract local template asset persistence and optimistic revision updates.
- `src/OfficeAgent.Core/Services/ITemplateCatalog.cs`
  Responsibility: abstract higher-level template workflows such as list/apply/save/save-as/detach for a worksheet.
- `src/OfficeAgent.Core/Templates/TemplateDefinitionNormalizer.cs`
  Responsibility: convert worksheet runtime metadata into sheet-name-agnostic template definitions and back.
- `src/OfficeAgent.Core/Templates/TemplateFingerprintBuilder.cs`
  Responsibility: deterministically fingerprint a normalized template snapshot while ignoring workbook-local fields such as `SheetName`.
- `src/OfficeAgent.Core/Templates/WorksheetTemplateCatalog.cs`
  Responsibility: orchestrate local template workflows by combining the connector registry, worksheet metadata stores, and template store.
- `src/OfficeAgent.Infrastructure/Storage/LocalJsonTemplateStore.cs`
  Responsibility: persist template assets under `%LocalAppData%\OfficeAgent\templates\<systemKey>\<projectId>\<templateId>.json`.
- `src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs`
  Responsibility: render and parse `TemplateBindings`, `SheetBindings`, and `SheetFieldMappings` in a fixed, readable order.
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
  Responsibility: round-trip `TemplateBindings` rows alongside the existing runtime metadata tables, including workbook-scope caching.
- `src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs`
  Responsibility: own template-related ribbon state and commands without mixing template logic into `RibbonSyncController`.
- `src/OfficeAgent.ExcelAddIn/Dialogs/TemplateDialogService.cs`
  Responsibility: host the template-dialog abstraction used by the controller and the concrete WinForms implementation.
- `src/OfficeAgent.ExcelAddIn/Dialogs/TemplatePickerDialog.cs`
  Responsibility: let users choose one template from the current project?s local catalog before applying it.
- `src/OfficeAgent.ExcelAddIn/Dialogs/TemplateNameDialog.cs`
  Responsibility: capture the name for ??????.
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Responsibility: bind template buttons to `RibbonTemplateController`, refresh enablement, and keep controller refreshes aligned with existing sync-controller events.
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
  Responsibility: declare the new template ribbon group and its three buttons.
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Responsibility: compose the new template store/catalog/controller and refresh them on workbook/sheet lifecycle events.
- `tests/OfficeAgent.Core.Tests/WorksheetTemplateCatalogTests.cs`
  Responsibility: lock normalization, dirty detection, apply/save/save-as, and revision conflict behavior at the orchestration layer.
- `tests/OfficeAgent.Infrastructure.Tests/LocalJsonTemplateStoreTests.cs`
  Responsibility: lock file layout, JSON round-trip, project filtering, and optimistic revision semantics.
- `tests/OfficeAgent.ExcelAddIn.Tests/MetadataSheetLayoutSerializerTests.cs`
  Responsibility: verify the new three-section `AI_Setting` layout.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`
  Responsibility: verify `TemplateBindings` round-trip, workbook-scoped cache invalidation, and sheet-specific clears.
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonTemplateControllerTests.cs`
  Responsibility: verify apply/save/save-as UX rules and controller state transitions without WinForms.
- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  Responsibility: regression-guard ribbon layout, button wiring, and controller-refresh plumbing by reading source text.
- `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: document the new `TemplateBindings` section and template actions.
- `docs/ribbon-sync-real-system-integration-guide.md`
  Responsibility: explain the new local-template layer and why runtime metadata is still the execution truth.
- `docs/vsto-manual-test-checklist.md`
  Responsibility: add manual steps for apply/save/save-as workflows and old-workbook compatibility.

### Task 1: Lock Template Catalog Contracts and Core Workflow with Tests

**Files:**
- Create: `tests/OfficeAgent.Core.Tests/WorksheetTemplateCatalogTests.cs`
- Create: `src/OfficeAgent.Core/Models/TemplateDefinition.cs`
- Create: `src/OfficeAgent.Core/Models/TemplateFieldMappingRow.cs`
- Create: `src/OfficeAgent.Core/Models/SheetTemplateBinding.cs`
- Create: `src/OfficeAgent.Core/Models/SheetTemplateState.cs`
- Create: `src/OfficeAgent.Core/Services/IWorksheetTemplateBindingStore.cs`
- Create: `src/OfficeAgent.Core/Services/ITemplateStore.cs`
- Create: `src/OfficeAgent.Core/Services/ITemplateCatalog.cs`
- Create: `src/OfficeAgent.Core/Templates/TemplateDefinitionNormalizer.cs`
- Create: `src/OfficeAgent.Core/Templates/TemplateFingerprintBuilder.cs`
- Create: `src/OfficeAgent.Core/Templates/WorksheetTemplateCatalog.cs`

- [ ] **Step 1: Write failing catalog tests for apply, save-as, dirty detection, and revision conflicts**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Templates;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class WorksheetTemplateCatalogTests
    {
        [Fact]
        public void ApplyTemplateToSheetWritesTemplateBindingAndInjectsSheetName()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "????");
            metadataStore.FieldMappings["Sheet1"] = CreateSheetMappings("Sheet1", "???");

            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            templateStore.Seed(CreateTemplate("tpl-performance-a", "??A", "???"));
            var registry = new SystemConnectorRegistry(new ISystemConnector[] { new FakeSystemConnector() });
            var catalog = new WorksheetTemplateCatalog(registry, metadataStore, templateBindingStore, templateStore);

            catalog.ApplyTemplateToSheet("Sheet1", "tpl-performance-a");

            Assert.Equal("??A", templateBindingStore.LastSavedBinding.TemplateName);
            Assert.Equal("store-template", templateBindingStore.LastSavedBinding.TemplateOrigin);
            Assert.Equal(1, templateBindingStore.LastSavedBinding.TemplateRevision);
            Assert.Equal("performance", metadataStore.LastSavedBinding.ProjectId);
            Assert.All(metadataStore.LastSavedFieldMappings, row => Assert.Equal("Sheet1", row.SheetName));
            Assert.Contains(metadataStore.LastSavedFieldMappings, row => row.Values["CurrentSingleDisplayName"] == "???");
        }

        [Fact]
        public void SaveSheetAsNewTemplateRemovesSheetNameAndBindsCurrentSheetToNewTemplate()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "????");
            metadataStore.FieldMappings["Sheet1"] = CreateSheetMappings("Sheet1", "?????");

            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            var templateStore = new FakeTemplateStore();
            var registry = new SystemConnectorRegistry(new ISystemConnector[] { new FakeSystemConnector() });
            var catalog = new WorksheetTemplateCatalog(registry, metadataStore, templateBindingStore, templateStore);

            catalog.SaveSheetAsNewTemplate("Sheet1", "??B");

            var savedTemplate = Assert.Single(templateStore.SavedNewTemplates);
            Assert.Equal("??B", savedTemplate.TemplateName);
            Assert.All(savedTemplate.FieldMappings, row => Assert.DoesNotContain("Sheet1", row.Values.Values));
            Assert.Equal("store-template", templateBindingStore.LastSavedBinding.TemplateOrigin);
            Assert.Equal("??B", templateBindingStore.LastSavedBinding.TemplateName);
            Assert.False(string.IsNullOrWhiteSpace(templateBindingStore.LastSavedBinding.TemplateId));
        }

        [Fact]
        public void GetSheetStateMarksSheetDirtyWhenCurrentFingerprintDiffersFromAppliedFingerprint()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "????");
            metadataStore.FieldMappings["Sheet1"] = CreateSheetMappings("Sheet1", "????");

            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            templateBindingStore.Bindings["Sheet1"] = new SheetTemplateBinding
            {
                SheetName = "Sheet1",
                TemplateId = "tpl-performance-a",
                TemplateName = "??A",
                TemplateRevision = 2,
                TemplateOrigin = "store-template",
                AppliedFingerprint = "stale-fingerprint",
            };

            var templateStore = new FakeTemplateStore();
            templateStore.Seed(CreateTemplate("tpl-performance-a", "??A", "???", revision: 2));
            var registry = new SystemConnectorRegistry(new ISystemConnector[] { new FakeSystemConnector() });
            var catalog = new WorksheetTemplateCatalog(registry, metadataStore, templateBindingStore, templateStore);

            var state = catalog.GetSheetState("Sheet1");

            Assert.True(state.HasProjectBinding);
            Assert.True(state.CanApplyTemplate);
            Assert.True(state.CanSaveTemplate);
            Assert.True(state.CanSaveAsTemplate);
            Assert.True(state.IsDirty);
            Assert.False(state.TemplateMissing);
        }

        [Fact]
        public void SaveSheetToExistingTemplateRejectsRevisionMismatchWithoutOverwrite()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "????");
            metadataStore.FieldMappings["Sheet1"] = CreateSheetMappings("Sheet1", "???");

            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            templateBindingStore.Bindings["Sheet1"] = new SheetTemplateBinding
            {
                SheetName = "Sheet1",
                TemplateId = "tpl-performance-a",
                TemplateName = "??A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
                AppliedFingerprint = "up-to-date",
            };

            var templateStore = new FakeTemplateStore();
            templateStore.Seed(CreateTemplate("tpl-performance-a", "??A", "???", revision: 3));
            var registry = new SystemConnectorRegistry(new ISystemConnector[] { new FakeSystemConnector() });
            var catalog = new WorksheetTemplateCatalog(registry, metadataStore, templateBindingStore, templateStore);

            var error = Assert.Throws<InvalidOperationException>(() =>
                catalog.SaveSheetToExistingTemplate("Sheet1", overwriteRevisionConflict: false));

            Assert.Equal("????????", error.Message);
        }

        [Fact]
        public void GetSheetStateMarksTemplateMissingWhenTemplateRecordNoLongerExists()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "????");
            metadataStore.FieldMappings["Sheet1"] = CreateSheetMappings("Sheet1", "???");

            var templateBindingStore = new FakeWorksheetTemplateBindingStore();
            templateBindingStore.Bindings["Sheet1"] = new SheetTemplateBinding
            {
                SheetName = "Sheet1",
                TemplateId = "tpl-missing",
                TemplateName = "?????",
                TemplateRevision = 2,
                TemplateOrigin = "store-template",
                AppliedFingerprint = "stale-fingerprint",
            };

            var templateStore = new FakeTemplateStore();
            var registry = new SystemConnectorRegistry(new ISystemConnector[] { new FakeSystemConnector() });
            var catalog = new WorksheetTemplateCatalog(registry, metadataStore, templateBindingStore, templateStore);

            var state = catalog.GetSheetState("Sheet1");

            Assert.True(state.HasProjectBinding);
            Assert.True(state.CanApplyTemplate);
            Assert.True(state.CanSaveAsTemplate);
            Assert.False(state.CanSaveTemplate);
            Assert.True(state.TemplateMissing);
            Assert.Equal("?????", state.TemplateName);
        }

        private static SheetBinding CreateBinding(string sheetName, string projectId, string projectName)
        {
            return new SheetBinding
            {
                SheetName = sheetName,
                SystemKey = "current-business-system",
                ProjectId = projectId,
                ProjectName = projectName,
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
        }

        private static SheetFieldMappingRow[] CreateSheetMappings(string sheetName, string currentHeaderText)
        {
            return new[]
            {
                new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["HeaderId"] = "owner_name",
                        ["ApiFieldKey"] = "owner_name",
                        ["CurrentSingleDisplayName"] = currentHeaderText,
                    },
                },
            };
        }

        private static TemplateDefinition CreateTemplate(string templateId, string templateName, string currentHeaderText, int revision = 1)
        {
            return new TemplateDefinition
            {
                TemplateId = templateId,
                TemplateName = templateName,
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "????",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = "current-business-system",
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderId },
                        new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
                        new FieldMappingColumnDefinition { ColumnName = "CurrentSingleDisplayName", Role = FieldMappingSemanticRole.CurrentSingleHeaderText },
                    },
                },
                FieldMappings = new[]
                {
                    new TemplateFieldMappingRow
                    {
                        Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["HeaderId"] = "owner_name",
                            ["ApiFieldKey"] = "owner_name",
                            ["CurrentSingleDisplayName"] = currentHeaderText,
                        },
                    },
                },
                Revision = revision,
                CreatedAtUtc = DateTime.UtcNow,
                UpdatedAtUtc = DateTime.UtcNow,
            };
        }

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public string SystemKey => "current-business-system";

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return new[]
                {
                    new ProjectOption
                    {
                        SystemKey = SystemKey,
                        ProjectId = "performance",
                        DisplayName = "????",
                    },
                };
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                return CreateBinding(sheetName, project?.ProjectId ?? "performance", project?.DisplayName ?? "????");
            }

            public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
            {
                return CreateTemplate("seed", "seed", "???").FieldMappingDefinition;
            }

            public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
            {
                return CreateSheetMappings(sheetName, "???");
            }

            public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
            {
                return Array.Empty<IDictionary<string, object>>();
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);
            public SheetBinding LastSavedBinding { get; private set; }
            public FieldMappingTableDefinition LastSavedFieldMappingDefinition { get; private set; }
            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public void SaveBinding(SheetBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public SheetBinding LoadBinding(string sheetName)
            {
                if (!Bindings.TryGetValue(sheetName, out var binding))
                {
                    throw new InvalidOperationException("Binding missing.");
                }

                return binding;
            }

            public void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows)
            {
                LastSavedFieldMappingDefinition = definition;
                LastSavedFieldMappings = (rows ?? Array.Empty<SheetFieldMappingRow>()).ToArray();
                FieldMappings[sheetName] = LastSavedFieldMappings;
            }

            public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition)
            {
                return FieldMappings.TryGetValue(sheetName, out var rows)
                    ? rows
                    : Array.Empty<SheetFieldMappingRow>();
            }

            public void ClearFieldMappings(string sheetName)
            {
                FieldMappings.Remove(sheetName);
            }

            public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
            {
                return Array.Empty<WorksheetSnapshotCell>();
            }

            public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
            {
            }
        }

        private sealed class FakeWorksheetTemplateBindingStore : IWorksheetTemplateBindingStore
        {
            public Dictionary<string, SheetTemplateBinding> Bindings { get; } = new Dictionary<string, SheetTemplateBinding>(StringComparer.OrdinalIgnoreCase);
            public SheetTemplateBinding LastSavedBinding { get; private set; }

            public void SaveTemplateBinding(SheetTemplateBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public SheetTemplateBinding LoadTemplateBinding(string sheetName)
            {
                if (!Bindings.TryGetValue(sheetName, out var binding))
                {
                    throw new InvalidOperationException("Template binding missing.");
                }

                return binding;
            }

            public void ClearTemplateBinding(string sheetName)
            {
                Bindings.Remove(sheetName);
            }
        }

        private sealed class FakeTemplateStore : ITemplateStore
        {
            private readonly Dictionary<string, TemplateDefinition> templates = new Dictionary<string, TemplateDefinition>(StringComparer.Ordinal);

            public List<TemplateDefinition> SavedNewTemplates { get; } = new List<TemplateDefinition>();

            public void Seed(TemplateDefinition template)
            {
                templates[template.TemplateId] = template;
            }

            public IReadOnlyList<TemplateDefinition> ListByProject(string systemKey, string projectId)
            {
                return templates.Values
                    .Where(template =>
                        string.Equals(template.SystemKey, systemKey, StringComparison.Ordinal) &&
                        string.Equals(template.ProjectId, projectId, StringComparison.Ordinal))
                    .ToArray();
            }

            public TemplateDefinition Load(string templateId)
            {
                if (!templates.TryGetValue(templateId, out var template))
                {
                    throw new InvalidOperationException("??????");
                }

                return template;
            }

            public TemplateDefinition SaveNew(TemplateDefinition template)
            {
                SavedNewTemplates.Add(template);
                templates[template.TemplateId] = template;
                return template;
            }

            public TemplateDefinition SaveExisting(TemplateDefinition template, int expectedRevision)
            {
                if (!templates.TryGetValue(template.TemplateId, out var existing))
                {
                    throw new InvalidOperationException("??????");
                }

                if (existing.Revision != expectedRevision)
                {
                    throw new InvalidOperationException("????????");
                }

                template.Revision = existing.Revision + 1;
                templates[template.TemplateId] = template;
                return template;
            }
        }
    }
}
```

- [ ] **Step 2: Run the targeted core tests to verify they fail for the intended reason**

Run:
```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~WorksheetTemplateCatalogTests"
```

Expected:
- FAIL with missing types such as `TemplateDefinition`, `SheetTemplateBinding`, `IWorksheetTemplateBindingStore`, or `WorksheetTemplateCatalog`

- [ ] **Step 3: Add the new models and service contracts that the tests exercise**

```csharp
using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class TemplateFieldMappingRow
    {
        public IReadOnlyDictionary<string, string> Values { get; set; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class TemplateDefinition
    {
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public int HeaderStartRow { get; set; } = 1;
        public int HeaderRowCount { get; set; } = 2;
        public int DataStartRow { get; set; } = 3;
        public FieldMappingTableDefinition FieldMappingDefinition { get; set; } = new FieldMappingTableDefinition();
        public string FieldMappingDefinitionFingerprint { get; set; } = string.Empty;
        public TemplateFieldMappingRow[] FieldMappings { get; set; } = Array.Empty<TemplateFieldMappingRow>();
        public int Revision { get; set; } = 1;
        public DateTime CreatedAtUtc { get; set; }
        public DateTime UpdatedAtUtc { get; set; }
    }

    public sealed class SheetTemplateBinding
    {
        public string SheetName { get; set; } = string.Empty;
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public int TemplateRevision { get; set; }
        public string TemplateOrigin { get; set; } = "ad-hoc";
        public string AppliedFingerprint { get; set; } = string.Empty;
        public string TemplateLastAppliedAt { get; set; } = string.Empty;
        public string DerivedFromTemplateId { get; set; } = string.Empty;
        public int DerivedFromTemplateRevision { get; set; }
    }

    public sealed class SheetTemplateState
    {
        public bool HasProjectBinding { get; set; }
        public bool CanApplyTemplate { get; set; }
        public bool CanSaveTemplate { get; set; }
        public bool CanSaveAsTemplate { get; set; }
        public bool IsDirty { get; set; }
        public bool TemplateMissing { get; set; }
        public string ProjectDisplayName { get; set; } = string.Empty;
        public string TemplateId { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public int TemplateRevision { get; set; }
        public int StoredTemplateRevision { get; set; }
        public string TemplateOrigin { get; set; } = "ad-hoc";
    }
}

namespace OfficeAgent.Core.Services
{
    public interface IWorksheetTemplateBindingStore
    {
        void SaveTemplateBinding(SheetTemplateBinding binding);
        SheetTemplateBinding LoadTemplateBinding(string sheetName);
        void ClearTemplateBinding(string sheetName);
    }

    public interface ITemplateStore
    {
        IReadOnlyList<TemplateDefinition> ListByProject(string systemKey, string projectId);
        TemplateDefinition Load(string templateId);
        TemplateDefinition SaveNew(TemplateDefinition template);
        TemplateDefinition SaveExisting(TemplateDefinition template, int expectedRevision);
    }

    public interface ITemplateCatalog
    {
        IReadOnlyList<TemplateDefinition> ListTemplates(string sheetName);
        SheetTemplateState GetSheetState(string sheetName);
        void ApplyTemplateToSheet(string sheetName, string templateId);
        void SaveSheetToExistingTemplate(string sheetName, bool overwriteRevisionConflict);
        void SaveSheetAsNewTemplate(string sheetName, string templateName);
        void DetachTemplate(string sheetName);
    }
}
```

- [ ] **Step 4: Implement the normalizer, fingerprint builder, and worksheet template catalog**

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Templates
{
    internal static class TemplateDefinitionNormalizer
    {
        public static TemplateDefinition FromWorksheet(
            SheetBinding binding,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> rows,
            string templateId,
            string templateName,
            int revision,
            DateTime createdAtUtc,
            DateTime updatedAtUtc)
        {
            return new TemplateDefinition
            {
                TemplateId = templateId ?? string.Empty,
                TemplateName = templateName ?? string.Empty,
                SystemKey = binding?.SystemKey ?? string.Empty,
                ProjectId = binding?.ProjectId ?? string.Empty,
                ProjectName = binding?.ProjectName ?? string.Empty,
                HeaderStartRow = binding?.HeaderStartRow ?? 1,
                HeaderRowCount = binding?.HeaderRowCount ?? 2,
                DataStartRow = binding?.DataStartRow ?? 3,
                FieldMappingDefinition = definition ?? new FieldMappingTableDefinition(),
                FieldMappings = (rows ?? Array.Empty<SheetFieldMappingRow>())
                    .Select(row => new TemplateFieldMappingRow
                    {
                        Values = (row.Values ?? new Dictionary<string, string>())
                            .ToDictionary(pair => pair.Key, pair => pair.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase),
                    })
                    .ToArray(),
                Revision = revision,
                CreatedAtUtc = createdAtUtc,
                UpdatedAtUtc = updatedAtUtc,
            };
        }

        public static SheetBinding ToSheetBinding(string sheetName, TemplateDefinition template)
        {
            return new SheetBinding
            {
                SheetName = sheetName ?? string.Empty,
                SystemKey = template?.SystemKey ?? string.Empty,
                ProjectId = template?.ProjectId ?? string.Empty,
                ProjectName = template?.ProjectName ?? string.Empty,
                HeaderStartRow = template?.HeaderStartRow ?? 1,
                HeaderRowCount = template?.HeaderRowCount ?? 2,
                DataStartRow = template?.DataStartRow ?? 3,
            };
        }

        public static SheetFieldMappingRow[] ToSheetMappings(string sheetName, IReadOnlyList<TemplateFieldMappingRow> rows)
        {
            return (rows ?? Array.Empty<TemplateFieldMappingRow>())
                .Select(row => new SheetFieldMappingRow
                {
                    SheetName = sheetName ?? string.Empty,
                    Values = (row.Values ?? new Dictionary<string, string>())
                        .ToDictionary(pair => pair.Key, pair => pair.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase),
                })
                .ToArray();
        }
    }

    internal static class TemplateFingerprintBuilder
    {
        public static string Build(TemplateDefinition template)
        {
            var builder = new StringBuilder();
            builder.Append(template.SystemKey).Append('|')
                .Append(template.ProjectId).Append('|')
                .Append(template.ProjectName).Append('|')
                .Append(template.HeaderStartRow.ToString(CultureInfo.InvariantCulture)).Append('|')
                .Append(template.HeaderRowCount.ToString(CultureInfo.InvariantCulture)).Append('|')
                .Append(template.DataStartRow.ToString(CultureInfo.InvariantCulture)).Append('|');

            foreach (var column in template.FieldMappingDefinition?.Columns ?? Array.Empty<FieldMappingColumnDefinition>())
            {
                builder.Append(column.ColumnName).Append(':')
                    .Append(column.Role.ToString()).Append(':')
                    .Append(column.RoleKey ?? string.Empty).Append('|');
            }

            foreach (var row in template.FieldMappings ?? Array.Empty<TemplateFieldMappingRow>())
            {
                foreach (var pair in (row.Values ?? new Dictionary<string, string>()).OrderBy(pair => pair.Key, StringComparer.Ordinal))
                {
                    builder.Append(pair.Key).Append('=').Append(pair.Value ?? string.Empty).Append('|');
                }

                builder.Append(";");
            }

            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(builder.ToString()));
                return BitConverter.ToString(bytes).Replace("-", string.Empty);
            }
        }
    }

    public sealed class WorksheetTemplateCatalog : ITemplateCatalog
    {
        private readonly ISystemConnectorRegistry connectorRegistry;
        private readonly IWorksheetMetadataStore metadataStore;
        private readonly IWorksheetTemplateBindingStore templateBindingStore;
        private readonly ITemplateStore templateStore;

        public WorksheetTemplateCatalog(
            ISystemConnectorRegistry connectorRegistry,
            IWorksheetMetadataStore metadataStore,
            IWorksheetTemplateBindingStore templateBindingStore,
            ITemplateStore templateStore)
        {
            this.connectorRegistry = connectorRegistry ?? throw new ArgumentNullException(nameof(connectorRegistry));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.templateBindingStore = templateBindingStore ?? throw new ArgumentNullException(nameof(templateBindingStore));
            this.templateStore = templateStore ?? throw new ArgumentNullException(nameof(templateStore));
        }

        public IReadOnlyList<TemplateDefinition> ListTemplates(string sheetName)
        {
            var binding = metadataStore.LoadBinding(sheetName);
            return templateStore.ListByProject(binding.SystemKey, binding.ProjectId);
        }

        public SheetTemplateState GetSheetState(string sheetName)
        {
            try
            {
                var binding = metadataStore.LoadBinding(sheetName);
                var templateBinding = TryLoadTemplateBinding(sheetName);
                var state = new SheetTemplateState
                {
                    HasProjectBinding = true,
                    CanApplyTemplate = true,
                    CanSaveAsTemplate = true,
                    ProjectDisplayName = string.IsNullOrWhiteSpace(binding.ProjectName)
                        ? binding.ProjectId
                        : binding.ProjectName,
                    TemplateId = templateBinding?.TemplateId ?? string.Empty,
                    TemplateName = templateBinding?.TemplateName ?? string.Empty,
                    TemplateRevision = templateBinding?.TemplateRevision ?? 0,
                    TemplateOrigin = templateBinding?.TemplateOrigin ?? "ad-hoc",
                };

                if (templateBinding == null || !string.Equals(templateBinding.TemplateOrigin, "store-template", StringComparison.Ordinal))
                {
                    return state;
                }

                TemplateDefinition template;
                try
                {
                    template = templateStore.Load(templateBinding.TemplateId);
                }
                catch (InvalidOperationException)
                {
                    state.TemplateMissing = true;
                    state.CanSaveTemplate = false;
                    return state;
                }

                var current = BuildWorksheetTemplate(sheetName, template.TemplateId, template.TemplateName, templateBinding.TemplateRevision, template.CreatedAtUtc, template.UpdatedAtUtc);
                var currentFingerprint = TemplateFingerprintBuilder.Build(current);
                state.CanSaveTemplate = string.Equals(binding.SystemKey, template.SystemKey, StringComparison.Ordinal) &&
                                        string.Equals(binding.ProjectId, template.ProjectId, StringComparison.Ordinal);
                state.IsDirty = !string.Equals(currentFingerprint, templateBinding.AppliedFingerprint, StringComparison.Ordinal);
                state.StoredTemplateRevision = template.Revision;
                state.TemplateMissing = false;
                return state;
            }
            catch (InvalidOperationException)
            {
                return new SheetTemplateState();
            }
        }

        public void ApplyTemplateToSheet(string sheetName, string templateId)
        {
            var template = templateStore.Load(templateId);
            var currentBinding = metadataStore.LoadBinding(sheetName);
            EnsureCompatible(template, currentBinding);
            metadataStore.SaveBinding(TemplateDefinitionNormalizer.ToSheetBinding(sheetName, template));
            metadataStore.SaveFieldMappings(sheetName, template.FieldMappingDefinition, TemplateDefinitionNormalizer.ToSheetMappings(sheetName, template.FieldMappings));
            templateBindingStore.SaveTemplateBinding(new SheetTemplateBinding
            {
                SheetName = sheetName,
                TemplateId = template.TemplateId,
                TemplateName = template.TemplateName,
                TemplateRevision = template.Revision,
                TemplateOrigin = "store-template",
                AppliedFingerprint = TemplateFingerprintBuilder.Build(BuildWorksheetTemplate(sheetName, template.TemplateId, template.TemplateName, template.Revision, template.CreatedAtUtc, template.UpdatedAtUtc)),
                TemplateLastAppliedAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            });
        }

        public void SaveSheetToExistingTemplate(string sheetName, bool overwriteRevisionConflict)
        {
            var binding = metadataStore.LoadBinding(sheetName);
            var templateBinding = templateBindingStore.LoadTemplateBinding(sheetName);
            if (!string.Equals(templateBinding.TemplateOrigin, "store-template", StringComparison.Ordinal) || string.IsNullOrWhiteSpace(templateBinding.TemplateId))
            {
                throw new InvalidOperationException("????????????");
            }

            var existing = templateStore.Load(templateBinding.TemplateId);
            if (!overwriteRevisionConflict && existing.Revision != templateBinding.TemplateRevision)
            {
                throw new InvalidOperationException("????????");
            }

            if (!string.Equals(binding.SystemKey, existing.SystemKey, StringComparison.Ordinal) ||
                !string.Equals(binding.ProjectId, existing.ProjectId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("??????????????");
            }

            var updated = BuildWorksheetTemplate(sheetName, existing.TemplateId, existing.TemplateName, existing.Revision, existing.CreatedAtUtc, DateTime.UtcNow);
            var saved = templateStore.SaveExisting(updated, overwriteRevisionConflict ? existing.Revision : templateBinding.TemplateRevision);
            templateBindingStore.SaveTemplateBinding(new SheetTemplateBinding
            {
                SheetName = sheetName,
                TemplateId = saved.TemplateId,
                TemplateName = saved.TemplateName,
                TemplateRevision = saved.Revision,
                TemplateOrigin = "store-template",
                AppliedFingerprint = TemplateFingerprintBuilder.Build(saved),
                TemplateLastAppliedAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
                DerivedFromTemplateId = templateBinding.DerivedFromTemplateId,
                DerivedFromTemplateRevision = templateBinding.DerivedFromTemplateRevision,
            });
        }

        public void SaveSheetAsNewTemplate(string sheetName, string templateName)
        {
            var existingBinding = TryLoadTemplateBinding(sheetName);
            var now = DateTime.UtcNow;
            var created = BuildWorksheetTemplate(sheetName, Guid.NewGuid().ToString("N"), templateName, revision: 1, createdAtUtc: now, updatedAtUtc: now);
            var saved = templateStore.SaveNew(created);
            templateBindingStore.SaveTemplateBinding(new SheetTemplateBinding
            {
                SheetName = sheetName,
                TemplateId = saved.TemplateId,
                TemplateName = saved.TemplateName,
                TemplateRevision = saved.Revision,
                TemplateOrigin = "store-template",
                AppliedFingerprint = TemplateFingerprintBuilder.Build(saved),
                TemplateLastAppliedAt = now.ToString("O", CultureInfo.InvariantCulture),
                DerivedFromTemplateId = existingBinding?.TemplateId ?? string.Empty,
                DerivedFromTemplateRevision = existingBinding?.TemplateRevision ?? 0,
            });
        }

        public void DetachTemplate(string sheetName)
        {
            templateBindingStore.SaveTemplateBinding(new SheetTemplateBinding
            {
                SheetName = sheetName,
                TemplateOrigin = "ad-hoc",
            });
        }

        private TemplateDefinition BuildWorksheetTemplate(string sheetName, string templateId, string templateName, int revision, DateTime createdAtUtc, DateTime updatedAtUtc)
        {
            var binding = metadataStore.LoadBinding(sheetName);
            var definition = connectorRegistry.GetRequiredConnector(binding.SystemKey).GetFieldMappingDefinition(binding.ProjectId);
            var rows = metadataStore.LoadFieldMappings(sheetName, definition) ?? Array.Empty<SheetFieldMappingRow>();
            var template = TemplateDefinitionNormalizer.FromWorksheet(binding, definition, rows, templateId, templateName, revision, createdAtUtc, updatedAtUtc);
            template.FieldMappingDefinitionFingerprint = TemplateFingerprintBuilder.Build(new TemplateDefinition
            {
                SystemKey = template.SystemKey,
                ProjectId = template.ProjectId,
                ProjectName = template.ProjectName,
                HeaderStartRow = template.HeaderStartRow,
                HeaderRowCount = template.HeaderRowCount,
                DataStartRow = template.DataStartRow,
                FieldMappingDefinition = template.FieldMappingDefinition,
                FieldMappings = Array.Empty<TemplateFieldMappingRow>(),
            });
            return template;
        }

        private static void EnsureCompatible(TemplateDefinition template, SheetBinding binding)
        {
            if (!string.Equals(template.SystemKey, binding.SystemKey, StringComparison.Ordinal) ||
                !string.Equals(template.ProjectId, binding.ProjectId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("???????????");
            }
        }

        private SheetTemplateBinding TryLoadTemplateBinding(string sheetName)
        {
            try
            {
                return templateBindingStore.LoadTemplateBinding(sheetName);
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }
    }
}
```

- [ ] **Step 5: Run the targeted core tests to verify they pass**

Run:
```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~WorksheetTemplateCatalogTests"
```

Expected:
- PASS

- [ ] **Step 6: Commit the core template contracts and orchestration**

```bash
git add -- tests/OfficeAgent.Core.Tests/WorksheetTemplateCatalogTests.cs src/OfficeAgent.Core/Models/TemplateDefinition.cs src/OfficeAgent.Core/Models/TemplateFieldMappingRow.cs src/OfficeAgent.Core/Models/SheetTemplateBinding.cs src/OfficeAgent.Core/Models/SheetTemplateState.cs src/OfficeAgent.Core/Services/IWorksheetTemplateBindingStore.cs src/OfficeAgent.Core/Services/ITemplateStore.cs src/OfficeAgent.Core/Services/ITemplateCatalog.cs src/OfficeAgent.Core/Templates/TemplateDefinitionNormalizer.cs src/OfficeAgent.Core/Templates/TemplateFingerprintBuilder.cs src/OfficeAgent.Core/Templates/WorksheetTemplateCatalog.cs
git commit -m "feat: add ribbon sync template catalog core"
```

### Task 2: Add `TemplateBindings` to `AI_Setting` Layout and Metadata Storage

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/MetadataSheetLayoutSerializerTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`

- [ ] **Step 1: Extend serializer and metadata-store tests to cover `TemplateBindings`**

```csharp
[Fact]
public void RenderPlacesTemplateBindingsAboveSheetBindingsAndFieldMappings()
{
    var serializer = CreateSerializer();
    var rendered = InvokeRender(
        serializer,
        new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
        {
            ["TemplateBindings"] = CreateSection(
                "TemplateBindings",
                new[] { "SheetName", "TemplateId", "TemplateOrigin" },
                new[] { new[] { "Sheet1", "tpl-performance-a", "store-template" } }),
            ["SheetBindings"] = CreateSection(
                "SheetBindings",
                new[] { "SheetName", "SystemKey" },
                new[] { new[] { "Sheet1", "current-business-system" } }),
            ["SheetFieldMappings"] = CreateSection(
                "SheetFieldMappings",
                new[] { "SheetName", "HeaderId", "ApiFieldKey" },
                new[] { new[] { "Sheet1", "row_id", "row_id" } }),
        });

    Assert.Equal("TemplateBindings", rendered[0][0]);
    Assert.Equal("SheetBindings", rendered[5][0]);
    Assert.Equal("SheetFieldMappings", rendered[10][0]);
}

[Fact]
public void SaveTemplateBindingRoundTripsTemplateMetadata()
{
    var (store, adapter) = CreateStore();
    var binding = new SheetTemplateBinding
    {
        SheetName = "Sheet1",
        TemplateId = "tpl-performance-a",
        TemplateName = "??A",
        TemplateRevision = 3,
        TemplateOrigin = "store-template",
        AppliedFingerprint = "ABC123",
        TemplateLastAppliedAt = "2026-04-22T08:00:00.0000000Z",
        DerivedFromTemplateId = "tpl-base",
        DerivedFromTemplateRevision = 1,
    };

    InvokeSaveTemplateBinding(store, binding);
    var loaded = InvokeLoadTemplateBinding(store, "Sheet1");

    Assert.Equal("tpl-performance-a", loaded.TemplateId);
    Assert.Equal("??A", loaded.TemplateName);
    Assert.Equal(3, loaded.TemplateRevision);
    Assert.Equal("store-template", loaded.TemplateOrigin);
    Assert.Equal("ABC123", loaded.AppliedFingerprint);
    Assert.Equal("tpl-base", loaded.DerivedFromTemplateId);
    Assert.Equal(1, loaded.DerivedFromTemplateRevision);
}

[Fact]
public void ClearTemplateBindingRemovesOnlyRequestedSheet()
{
    var (store, adapter) = CreateStore();
    adapter.SeedTable("TemplateBindings", new[]
    {
        new[] { "Sheet1", "tpl-a", "??A", "1", "store-template", "A1", "2026-04-22T08:00:00.0000000Z", "", "0" },
        new[] { "Sheet2", "tpl-b", "??B", "2", "store-template", "B2", "2026-04-22T09:00:00.0000000Z", "", "0" },
    });

    InvokeClearTemplateBinding(store, "Sheet1");

    var rows = adapter.ReadSeededTable("TemplateBindings");
    Assert.Single(rows);
    Assert.Equal("Sheet2", rows[0][0]);
}

private static void InvokeSaveTemplateBinding(object store, SheetTemplateBinding binding)
{
    var method = store.GetType().GetMethod("SaveTemplateBinding", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
    method.Invoke(store, new object[] { binding });
}

private static SheetTemplateBinding InvokeLoadTemplateBinding(object store, string sheetName)
{
    var method = store.GetType().GetMethod("LoadTemplateBinding", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
    return (SheetTemplateBinding)method.Invoke(store, new object[] { sheetName });
}

private static void InvokeClearTemplateBinding(object store, string sheetName)
{
    var method = store.GetType().GetMethod("ClearTemplateBinding", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
    method.Invoke(store, new object[] { sheetName });
}
```

- [ ] **Step 2: Run the targeted Excel add-in tests to verify they fail**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~MetadataSheetLayoutSerializerTests|FullyQualifiedName~WorksheetMetadataStoreTests"
```

Expected:
- FAIL because `TemplateBindings` is not part of section order and `WorksheetMetadataStore` does not yet implement the new template-binding methods

- [ ] **Step 3: Add `TemplateBindings` support to the serializer and metadata store**

```csharp
using System.Globalization;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetMetadataStore : IWorksheetMetadataStore, IWorksheetTemplateBindingStore
    {
        private const string TemplateBindingsTableName = "TemplateBindings";

        private static readonly string[] TemplateBindingHeaders =
        {
            "SheetName",
            "TemplateId",
            "TemplateName",
            "TemplateRevision",
            "TemplateOrigin",
            "AppliedFingerprint",
            "TemplateLastAppliedAt",
            "DerivedFromTemplateId",
            "DerivedFromTemplateRevision",
        };

        private string[][] templateBindingRowsCache;
        private bool templateBindingRowsCacheLoaded;

        public void SaveTemplateBinding(SheetTemplateBinding binding)
        {
            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (string.IsNullOrWhiteSpace(binding.SheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(binding));
            }

            EnsureWorkbookScope();
            adapter.EnsureWorksheet(MetadataSheetName, visible: true);
            var rows = GetTemplateBindingRows().ToList();
            var newRow = new[]
            {
                binding.SheetName,
                binding.TemplateId ?? string.Empty,
                binding.TemplateName ?? string.Empty,
                binding.TemplateRevision.ToString(CultureInfo.InvariantCulture),
                binding.TemplateOrigin ?? "ad-hoc",
                binding.AppliedFingerprint ?? string.Empty,
                binding.TemplateLastAppliedAt ?? string.Empty,
                binding.DerivedFromTemplateId ?? string.Empty,
                binding.DerivedFromTemplateRevision.ToString(CultureInfo.InvariantCulture),
            };

            var existingIndex = rows.FindIndex(row =>
                row.Length > 0 &&
                string.Equals(row[0], binding.SheetName, StringComparison.OrdinalIgnoreCase));

            if (existingIndex >= 0)
            {
                rows[existingIndex] = newRow;
            }
            else
            {
                rows.Add(newRow);
            }

            adapter.WriteTable(TemplateBindingsTableName, TemplateBindingHeaders, rows.ToArray());
            templateBindingRowsCache = CloneRows(rows);
            templateBindingRowsCacheLoaded = true;
        }

        public SheetTemplateBinding LoadTemplateBinding(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            EnsureWorkbookScope();
            var binding = GetTemplateBindingRows()
                .Select(ParseTemplateBindingRow)
                .FirstOrDefault(candidate =>
                    candidate != null &&
                    string.Equals(candidate.SheetName, sheetName, StringComparison.OrdinalIgnoreCase));

            if (binding != null)
            {
                return binding;
            }

            throw new InvalidOperationException($"Template binding for worksheet '{sheetName}' does not exist.");
        }

        public void ClearTemplateBinding(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            EnsureWorkbookScope();
            var rows = GetTemplateBindingRows().ToList();
            var removed = rows.RemoveAll(row =>
                row.Length > 0 &&
                string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

            if (removed == 0)
            {
                return;
            }

            adapter.WriteTable(TemplateBindingsTableName, TemplateBindingHeaders, rows.ToArray());
            templateBindingRowsCache = CloneRows(rows);
            templateBindingRowsCacheLoaded = true;
        }

        internal void InvalidateCache()
        {
            bindingRowsCache = null;
            bindingRowsCacheLoaded = false;
            templateBindingRowsCache = null;
            templateBindingRowsCacheLoaded = false;
            fieldMappingRowsCache = null;
            fieldMappingRowsCacheLoaded = false;
            fieldMappingHeaders = DefaultFieldMappingHeaders.ToArray();
        }

        private static SheetTemplateBinding ParseTemplateBindingRow(IReadOnlyList<string> row)
        {
            if (row == null || row.Count == 0 || string.IsNullOrWhiteSpace(row[0]))
            {
                return null;
            }

            return new SheetTemplateBinding
            {
                SheetName = row[0],
                TemplateId = row.Count > 1 ? row[1] : string.Empty,
                TemplateName = row.Count > 2 ? row[2] : string.Empty,
                TemplateRevision = ParseIntOrDefault(row, 3, 0),
                TemplateOrigin = row.Count > 4 ? row[4] : "ad-hoc",
                AppliedFingerprint = row.Count > 5 ? row[5] : string.Empty,
                TemplateLastAppliedAt = row.Count > 6 ? row[6] : string.Empty,
                DerivedFromTemplateId = row.Count > 7 ? row[7] : string.Empty,
                DerivedFromTemplateRevision = ParseIntOrDefault(row, 8, 0),
            };
        }

        private IReadOnlyList<string[]> GetTemplateBindingRows()
        {
            if (!templateBindingRowsCacheLoaded)
            {
                templateBindingRowsCache = adapter.ReadTable(TemplateBindingsTableName) ?? Array.Empty<string[]>();
                templateBindingRowsCacheLoaded = true;
            }

            return templateBindingRowsCache ?? Array.Empty<string[]>();
        }
    }
}

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class MetadataSheetLayoutSerializer
    {
        private static readonly string[] SectionOrder =
        {
            "TemplateBindings",
            "SheetBindings",
            "SheetFieldMappings",
        };
    }
}
```

- [ ] **Step 4: Run the targeted Excel add-in tests to verify they pass**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~MetadataSheetLayoutSerializerTests|FullyQualifiedName~WorksheetMetadataStoreTests"
```

Expected:
- PASS

- [ ] **Step 5: Commit the worksheet metadata changes**

```bash
git add -- src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs tests/OfficeAgent.ExcelAddIn.Tests/MetadataSheetLayoutSerializerTests.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs
git commit -m "feat: add ribbon sync template bindings metadata"
```

### Task 3: Implement the Local JSON Template Store

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Storage/LocalJsonTemplateStore.cs`
- Create: `tests/OfficeAgent.Infrastructure.Tests/LocalJsonTemplateStoreTests.cs`

- [ ] **Step 1: Write failing infrastructure tests for JSON round-trip, per-project listing, and revision checks**

```csharp
using System;
using System.IO;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class LocalJsonTemplateStoreTests : IDisposable
    {
        private readonly string rootPath = Path.Combine(Path.GetTempPath(), "OfficeAgent.TemplateStore.Tests", Guid.NewGuid().ToString("N"));

        [Fact]
        public void SaveNewAndLoadRoundTripTemplateDefinition()
        {
            var store = new LocalJsonTemplateStore(rootPath);
            var template = CreateTemplate("tpl-performance-a", "??A", revision: 1);

            var saved = store.SaveNew(template);
            var loaded = store.Load(saved.TemplateId);

            Assert.Equal("??A", loaded.TemplateName);
            Assert.Equal("current-business-system", loaded.SystemKey);
            Assert.Equal("performance", loaded.ProjectId);
            Assert.Single(loaded.FieldMappings);
            Assert.Equal("???", loaded.FieldMappings[0].Values["CurrentSingleDisplayName"]);
        }

        [Fact]
        public void ListByProjectReturnsOnlyMatchingProjectTemplatesOrderedByUpdatedAtDescending()
        {
            var store = new LocalJsonTemplateStore(rootPath);
            store.SaveNew(CreateTemplate("tpl-old", "???", revision: 1, updatedAtUtc: new DateTime(2026, 4, 22, 8, 0, 0, DateTimeKind.Utc)));
            store.SaveNew(CreateTemplate("tpl-new", "???", revision: 1, updatedAtUtc: new DateTime(2026, 4, 22, 9, 0, 0, DateTimeKind.Utc)));
            store.SaveNew(CreateTemplate("tpl-other", "????", revision: 1, projectId: "delivery-tracker", updatedAtUtc: new DateTime(2026, 4, 22, 10, 0, 0, DateTimeKind.Utc)));

            var templates = store.ListByProject("current-business-system", "performance");

            Assert.Equal(new[] { "tpl-new", "tpl-old" }, templates.Select(template => template.TemplateId).ToArray());
        }

        [Fact]
        public void SaveExistingRejectsRevisionMismatch()
        {
            var store = new LocalJsonTemplateStore(rootPath);
            store.SaveNew(CreateTemplate("tpl-performance-a", "??A", revision: 1));

            var error = Assert.Throws<InvalidOperationException>(() =>
                store.SaveExisting(CreateTemplate("tpl-performance-a", "??A-??", revision: 1), expectedRevision: 0));

            Assert.Equal("????????", error.Message);
        }

        private static TemplateDefinition CreateTemplate(string templateId, string templateName, int revision, string projectId = "performance", DateTime? updatedAtUtc = null)
        {
            return new TemplateDefinition
            {
                TemplateId = templateId,
                TemplateName = templateName,
                SystemKey = "current-business-system",
                ProjectId = projectId,
                ProjectName = projectId == "performance" ? "????" : "????",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
                FieldMappingDefinition = new FieldMappingTableDefinition(),
                FieldMappings = new[]
                {
                    new TemplateFieldMappingRow
                    {
                        Values = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["CurrentSingleDisplayName"] = "???",
                        },
                    },
                },
                Revision = revision,
                CreatedAtUtc = new DateTime(2026, 4, 22, 7, 0, 0, DateTimeKind.Utc),
                UpdatedAtUtc = updatedAtUtc ?? new DateTime(2026, 4, 22, 7, 0, 0, DateTimeKind.Utc),
            };
        }

        public void Dispose()
        {
            if (Directory.Exists(rootPath))
            {
                Directory.Delete(rootPath, recursive: true);
            }
        }
    }
}
```

- [ ] **Step 2: Run the targeted infrastructure tests to verify they fail**

Run:
```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter "FullyQualifiedName~LocalJsonTemplateStoreTests"
```

Expected:
- FAIL with `LocalJsonTemplateStore` missing

- [ ] **Step 3: Implement `LocalJsonTemplateStore` using the existing JSON storage pattern**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Infrastructure.Storage
{
    public sealed class LocalJsonTemplateStore : ITemplateStore
    {
        private readonly string rootPath;

        public LocalJsonTemplateStore(string rootPath)
        {
            this.rootPath = rootPath ?? throw new ArgumentNullException(nameof(rootPath));
        }

        public IReadOnlyList<TemplateDefinition> ListByProject(string systemKey, string projectId)
        {
            var directory = GetProjectDirectory(systemKey, projectId);
            if (!Directory.Exists(directory))
            {
                return Array.Empty<TemplateDefinition>();
            }

            return Directory.GetFiles(directory, "*.json", SearchOption.TopDirectoryOnly)
                .Select(path => ReadTemplate(path))
                .OrderByDescending(template => template.UpdatedAtUtc)
                .ToArray();
        }

        public TemplateDefinition Load(string templateId)
        {
            foreach (var file in Directory.Exists(rootPath)
                ? Directory.GetFiles(rootPath, "*.json", SearchOption.AllDirectories)
                : Array.Empty<string>())
            {
                var template = ReadTemplate(file);
                if (string.Equals(template.TemplateId, templateId, StringComparison.Ordinal))
                {
                    return template;
                }
            }

            throw new InvalidOperationException("??????");
        }

        public TemplateDefinition SaveNew(TemplateDefinition template)
        {
            var normalized = Normalize(template);
            normalized.Revision = Math.Max(1, normalized.Revision);
            if (normalized.CreatedAtUtc == default(DateTime))
            {
                normalized.CreatedAtUtc = DateTime.UtcNow;
            }

            if (normalized.UpdatedAtUtc == default(DateTime))
            {
                normalized.UpdatedAtUtc = normalized.CreatedAtUtc;
            }

            WriteTemplate(normalized);
            return Clone(normalized);
        }

        public TemplateDefinition SaveExisting(TemplateDefinition template, int expectedRevision)
        {
            var existing = Load(template.TemplateId);
            if (existing.Revision != expectedRevision)
            {
                throw new InvalidOperationException("????????");
            }

            var normalized = Normalize(template);
            normalized.CreatedAtUtc = existing.CreatedAtUtc;
            normalized.UpdatedAtUtc = DateTime.UtcNow;
            normalized.Revision = existing.Revision + 1;
            WriteTemplate(normalized);
            return Clone(normalized);
        }

        private void WriteTemplate(TemplateDefinition template)
        {
            var directory = GetProjectDirectory(template.SystemKey, template.ProjectId);
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, template.TemplateId + ".json");
            File.WriteAllText(path, JsonConvert.SerializeObject(template, Formatting.Indented));
        }

        private TemplateDefinition ReadTemplate(string path)
        {
            return Normalize(JsonConvert.DeserializeObject<TemplateDefinition>(File.ReadAllText(path)) ?? new TemplateDefinition());
        }

        private string GetProjectDirectory(string systemKey, string projectId)
        {
            return Path.Combine(rootPath, systemKey ?? string.Empty, projectId ?? string.Empty);
        }

        private static TemplateDefinition Normalize(TemplateDefinition template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            template.TemplateId = template.TemplateId ?? string.Empty;
            template.TemplateName = template.TemplateName ?? string.Empty;
            template.SystemKey = template.SystemKey ?? string.Empty;
            template.ProjectId = template.ProjectId ?? string.Empty;
            template.ProjectName = template.ProjectName ?? string.Empty;
            template.FieldMappingDefinition = template.FieldMappingDefinition ?? new FieldMappingTableDefinition();
            template.FieldMappings = template.FieldMappings ?? Array.Empty<TemplateFieldMappingRow>();
            return template;
        }

        private static TemplateDefinition Clone(TemplateDefinition template)
        {
            return JsonConvert.DeserializeObject<TemplateDefinition>(
                JsonConvert.SerializeObject(template)) ?? new TemplateDefinition();
        }
    }
}
```

- [ ] **Step 4: Run the targeted infrastructure tests to verify they pass**

Run:
```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter "FullyQualifiedName~LocalJsonTemplateStoreTests"
```

Expected:
- PASS

- [ ] **Step 5: Commit the local template store**

```bash
git add -- src/OfficeAgent.Infrastructure/Storage/LocalJsonTemplateStore.cs tests/OfficeAgent.Infrastructure.Tests/LocalJsonTemplateStoreTests.cs
git commit -m "feat: store ribbon sync templates locally"
```

### Task 4: Add a Dedicated Ribbon Template Controller and Test Its UX Rules

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/TemplateDialogService.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonTemplateControllerTests.cs`

- [ ] **Step 1: Write failing controller tests for apply, save, save-as, and conflict prompts**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class RibbonTemplateControllerTests
    {
        [Fact]
        public void RefreshActiveTemplateStateWithAdHocSheetEnablesApplyAndSaveAsButDisablesSave()
        {
            var catalog = new FakeTemplateCatalog
            {
                StateBySheet["Sheet1"] = new SheetTemplateState
                {
                    HasProjectBinding = true,
                    CanApplyTemplate = true,
                    CanSaveAsTemplate = true,
                    CanSaveTemplate = false,
                    TemplateOrigin = "ad-hoc",
                },
            };
            var dialogs = new FakeTemplateDialogService();
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeRefresh(controller);

            Assert.True(ReadBoolProperty(controller, "CanApplyTemplate"));
            Assert.True(ReadBoolProperty(controller, "CanSaveAsTemplate"));
            Assert.False(ReadBoolProperty(controller, "CanSaveTemplate"));
            Assert.Equal("?????", ReadStringProperty(controller, "ActiveTemplateDisplayName"));
        }

        [Fact]
        public void ExecuteApplyTemplateUsesSelectedTemplateAndRefreshesState()
        {
            var catalog = new FakeTemplateCatalog();
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanApplyTemplate = true,
                CanSaveAsTemplate = true,
            };
            catalog.TemplatesBySheet["Sheet1"] = new[]
            {
                new TemplateDefinition { TemplateId = "tpl-performance-a", TemplateName = "??A" },
            };
            var dialogs = new FakeTemplateDialogService
            {
                SelectedTemplateId = "tpl-performance-a",
            };
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeExecuteApplyTemplate(controller);

            Assert.Equal("tpl-performance-a", catalog.LastAppliedTemplateId);
            Assert.Contains("??A", dialogs.InfoMessages[0]);
        }

        [Fact]
        public void ExecuteSaveTemplateRoutesRevisionConflictToOverwriteOrSaveAsDialog()
        {
            var catalog = new FakeTemplateCatalog
            {
                ThrowRevisionConflictOnSave = true,
            };
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanSaveTemplate = true,
                CanSaveAsTemplate = true,
                TemplateId = "tpl-performance-a",
                TemplateName = "??A",
                TemplateRevision = 1,
                TemplateOrigin = "store-template",
            };
            var dialogs = new FakeTemplateDialogService
            {
                RevisionConflictResult = System.Windows.Forms.DialogResult.No,
                SaveAsTemplateName = "??B",
            };
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeExecuteSaveTemplate(controller);

            Assert.True(catalog.SaveExistingCalled);
            Assert.Equal("??B", catalog.LastSavedAsTemplateName);
        }

        private static object CreateController(
            ITemplateCatalog catalog,
            OfficeAgent.ExcelAddIn.Dialogs.IRibbonTemplateDialogService dialogs,
            Func<string> activeSheetNameProvider)
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var type = assembly.GetType("OfficeAgent.ExcelAddIn.RibbonTemplateController", throwOnError: true);
            var ctor = type.GetConstructor(
                BindingFlags.Instance | BindingFlags.NonPublic,
                binder: null,
                types: new[]
                {
                    typeof(ITemplateCatalog),
                    typeof(Func<string>),
                    typeof(OfficeAgent.ExcelAddIn.Dialogs.IRibbonTemplateDialogService),
                },
                modifiers: null);
            return ctor.Invoke(new object[] { catalog, activeSheetNameProvider, dialogs });
        }

        private static void InvokeRefresh(object controller)
        {
            var method = controller.GetType().GetMethod(
                "RefreshActiveTemplateStateFromSheetMetadata",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(controller, parameters: null);
        }

        private static void InvokeExecuteApplyTemplate(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteApplyTemplate",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(controller, parameters: null);
        }

        private static void InvokeExecuteSaveTemplate(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteSaveTemplate",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(controller, parameters: null);
        }

        private static bool ReadBoolProperty(object controller, string propertyName)
        {
            return (bool)controller.GetType()
                .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .GetValue(controller);
        }

        private static string ReadStringProperty(object controller, string propertyName)
        {
            return (string)controller.GetType()
                .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .GetValue(controller);
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }

        private sealed class FakeTemplateCatalog : ITemplateCatalog
        {
            public Dictionary<string, SheetTemplateState> StateBySheet { get; } = new Dictionary<string, SheetTemplateState>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, IReadOnlyList<TemplateDefinition>> TemplatesBySheet { get; } = new Dictionary<string, IReadOnlyList<TemplateDefinition>>(StringComparer.OrdinalIgnoreCase);
            public string LastAppliedTemplateId { get; private set; } = string.Empty;
            public string LastSavedAsTemplateName { get; private set; } = string.Empty;
            public bool SaveExistingCalled { get; private set; }
            public bool ThrowRevisionConflictOnSave { get; set; }

            public IReadOnlyList<TemplateDefinition> ListTemplates(string sheetName)
            {
                return TemplatesBySheet.TryGetValue(sheetName, out var templates)
                    ? templates
                    : Array.Empty<TemplateDefinition>();
            }

            public SheetTemplateState GetSheetState(string sheetName)
            {
                return StateBySheet.TryGetValue(sheetName, out var state)
                    ? state
                    : new SheetTemplateState();
            }

            public void ApplyTemplateToSheet(string sheetName, string templateId)
            {
                LastAppliedTemplateId = templateId;
            }

            public void SaveSheetToExistingTemplate(string sheetName, bool overwriteRevisionConflict)
            {
                SaveExistingCalled = true;
                if (ThrowRevisionConflictOnSave && !overwriteRevisionConflict)
                {
                    throw new InvalidOperationException("????????");
                }
            }

            public void SaveSheetAsNewTemplate(string sheetName, string templateName)
            {
                LastSavedAsTemplateName = templateName;
            }

            public void DetachTemplate(string sheetName)
            {
            }
        }

        private sealed class FakeTemplateDialogService : OfficeAgent.ExcelAddIn.Dialogs.IRibbonTemplateDialogService
        {
            public List<string> InfoMessages { get; } = new List<string>();
            public string SelectedTemplateId { get; set; } = string.Empty;
            public string SaveAsTemplateName { get; set; } = string.Empty;
            public System.Windows.Forms.DialogResult RevisionConflictResult { get; set; } = System.Windows.Forms.DialogResult.Cancel;

            public string ShowTemplatePicker(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates)
            {
                return SelectedTemplateId;
            }

            public string ShowSaveAsTemplateDialog(string suggestedTemplateName)
            {
                return SaveAsTemplateName;
            }

            public bool ConfirmApplyTemplateOverwrite(string templateName)
            {
                return true;
            }

            public System.Windows.Forms.DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision)
            {
                return RevisionConflictResult;
            }

            public void ShowInfo(string message)
            {
                InfoMessages.Add(message);
            }

            public void ShowWarning(string message)
            {
            }

            public void ShowError(string message)
            {
                throw new InvalidOperationException(message);
            }
        }
    }
}
```

- [ ] **Step 2: Run the targeted controller tests to verify they fail**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~RibbonTemplateControllerTests"
```

Expected:
- FAIL because `RibbonTemplateController` and `TemplateDialogService` do not exist yet

- [ ] **Step 3: Implement the controller and the dialog-service abstraction**

```csharp
using System;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    public interface IRibbonTemplateDialogService
    {
        string ShowTemplatePicker(string projectDisplayName, System.Collections.Generic.IReadOnlyList<TemplateDefinition> templates);
        string ShowSaveAsTemplateDialog(string suggestedTemplateName);
        bool ConfirmApplyTemplateOverwrite(string templateName);
        DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision);
        void ShowInfo(string message);
        void ShowWarning(string message);
        void ShowError(string message);
    }

    internal sealed class RibbonTemplateDialogService : IRibbonTemplateDialogService
    {
        public string ShowTemplatePicker(string projectDisplayName, System.Collections.Generic.IReadOnlyList<TemplateDefinition> templates)
        {
            using (var dialog = new TemplatePickerDialog(projectDisplayName, templates))
            {
                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.SelectedTemplateId
                    : string.Empty;
            }
        }

        public string ShowSaveAsTemplateDialog(string suggestedTemplateName)
        {
            using (var dialog = new TemplateNameDialog(suggestedTemplateName))
            {
                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.TemplateName
                    : string.Empty;
            }
        }

        public bool ConfirmApplyTemplateOverwrite(string templateName)
        {
            return MessageBox.Show(
                    $"????????????????????{templateName}?????",
                    "ISDP",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) == DialogResult.Yes;
        }

        public DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision)
        {
            return MessageBox.Show(
                $"???{templateName}????? {sheetRevision} ????? {storedRevision}?\r\n???????\r\n????????\r\n???????",
                "ISDP",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);
        }

        public void ShowInfo(string message) => OperationResultDialog.ShowInfo(message);
        public void ShowWarning(string message) => OperationResultDialog.ShowWarning(message);
        public void ShowError(string message) => OperationResultDialog.ShowError(message);
    }
}

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class RibbonTemplateController
    {
        private readonly ITemplateCatalog templateCatalog;
        private readonly Func<string> activeSheetNameProvider;
        private readonly Dialogs.IRibbonTemplateDialogService dialogService;
        private string lastRefreshedSheetName = string.Empty;

        public RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider)
            : this(templateCatalog, activeSheetNameProvider, new Dialogs.RibbonTemplateDialogService())
        {
        }

        internal RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider,
            Dialogs.IRibbonTemplateDialogService dialogService)
        {
            this.templateCatalog = templateCatalog ?? throw new ArgumentNullException(nameof(templateCatalog));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));
            ActiveTemplateDisplayName = "?????";
        }

        public event EventHandler TemplateStateChanged;

        public string ActiveTemplateDisplayName { get; private set; }
        public bool CanApplyTemplate { get; private set; }
        public bool CanSaveTemplate { get; private set; }
        public bool CanSaveAsTemplate { get; private set; }

        public void RefreshActiveTemplateStateFromSheetMetadata()
        {
            RefreshTemplateState(activeSheetNameProvider.Invoke() ?? string.Empty);
        }

        internal void RefreshTemplateState(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                ApplyState(new SheetTemplateState());
                lastRefreshedSheetName = string.Empty;
                return;
            }

            if (string.Equals(lastRefreshedSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            lastRefreshedSheetName = sheetName;
            ApplyState(templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState());
        }

        internal void InvalidateRefreshState()
        {
            lastRefreshedSheetName = string.Empty;
        }

        public void ExecuteApplyTemplate()
        {
            var sheetName = GetRequiredSheetName();
            var state = templateCatalog.GetSheetState(sheetName);
            if (!state.CanApplyTemplate)
            {
                dialogService.ShowWarning("???????");
                return;
            }

            var templates = templateCatalog.ListTemplates(sheetName);
            if (templates == null || templates.Count == 0)
            {
                dialogService.ShowWarning("???????????");
                return;
            }

            var templateId = dialogService.ShowTemplatePicker(state.ProjectDisplayName, templates);
            if (string.IsNullOrWhiteSpace(templateId))
            {
                return;
            }

            var selectedTemplate = templates.First(template => string.Equals(template.TemplateId, templateId, StringComparison.Ordinal));
            if (state.IsDirty && !dialogService.ConfirmApplyTemplateOverwrite(selectedTemplate.TemplateName))
            {
                return;
            }

            try
            {
                templateCatalog.ApplyTemplateToSheet(sheetName, templateId);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo($"???????\r\n???{selectedTemplate.TemplateName}");
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveTemplate()
        {
            var sheetName = GetRequiredSheetName();
            var state = templateCatalog.GetSheetState(sheetName);
            if (!state.CanSaveTemplate)
            {
                dialogService.ShowWarning("????????????");
                return;
            }

            try
            {
                templateCatalog.SaveSheetToExistingTemplate(sheetName, overwriteRevisionConflict: false);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo($"???????\r\n???{state.TemplateName}");
            }
            catch (InvalidOperationException ex) when (string.Equals(ex.Message, "????????", StringComparison.Ordinal))
            {
                var conflictResult = dialogService.ShowTemplateRevisionConflictDialog(state.TemplateName, state.TemplateRevision, state.StoredTemplateRevision);
                if (conflictResult == DialogResult.Yes)
                {
                    templateCatalog.SaveSheetToExistingTemplate(sheetName, overwriteRevisionConflict: true);
                    InvalidateRefreshState();
                    RefreshActiveTemplateStateFromSheetMetadata();
                    dialogService.ShowInfo($"???????\r\n???{state.TemplateName}");
                    return;
                }

                if (conflictResult == DialogResult.No)
                {
                    ExecuteSaveAsTemplate();
                    return;
                }
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveAsTemplate()
        {
            var sheetName = GetRequiredSheetName();
            var state = templateCatalog.GetSheetState(sheetName);
            if (!state.CanSaveAsTemplate)
            {
                dialogService.ShowWarning("???????");
                return;
            }

            var templateName = dialogService.ShowSaveAsTemplateDialog(string.IsNullOrWhiteSpace(state.TemplateName) ? "???" : state.TemplateName + "-??");
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return;
            }

            try
            {
                templateCatalog.SaveSheetAsNewTemplate(sheetName, templateName);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo($"???????\r\n???{templateName}");
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        private void ApplyState(SheetTemplateState state)
        {
            CanApplyTemplate = state?.CanApplyTemplate == true;
            CanSaveTemplate = state?.CanSaveTemplate == true;
            CanSaveAsTemplate = state?.CanSaveAsTemplate == true;
            ActiveTemplateDisplayName = string.IsNullOrWhiteSpace(state?.TemplateName)
                ? "?????"
                : state.TemplateName;
            TemplateStateChanged?.Invoke(this, EventArgs.Empty);
        }

        private string GetRequiredSheetName()
        {
            var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new InvalidOperationException("Active worksheet is not available.");
            }

            return sheetName;
        }
    }
}
```

- [ ] **Step 4: Run the targeted controller tests to verify they pass**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~RibbonTemplateControllerTests"
```

Expected:
- PASS

- [ ] **Step 5: Commit the ribbon template controller**

```bash
git add -- src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs src/OfficeAgent.ExcelAddIn/Dialogs/TemplateDialogService.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonTemplateControllerTests.cs
git commit -m "feat: add ribbon template controller"
```

### Task 5: Implement the Concrete Template Dialogs and Compose the Feature at Startup

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/TemplatePickerDialog.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/TemplateNameDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

- [ ] **Step 1: Add a simple picker dialog for per-project local templates**

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplatePickerDialog : Form
    {
        private readonly IReadOnlyList<TemplateDefinition> templates;
        private readonly ListBox templateListBox;

        public TemplatePickerDialog(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates)
        {
            this.templates = templates ?? Array.Empty<TemplateDefinition>();

            Text = "????";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(480, 320);

            var titleLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 48,
                Padding = new Padding(12, 12, 12, 0),
                Text = $"?????{projectDisplayName}\r\n????????????????",
            };

            templateListBox = new ListBox
            {
                Dock = DockStyle.Fill,
            };
            foreach (var template in this.templates.OrderByDescending(item => item.UpdatedAtUtc))
            {
                templateListBox.Items.Add(new TemplateListItem(template));
            }

            var okButton = new Button { Text = "??", Width = 88, Height = 30 };
            okButton.Click += OkButton_Click;
            var cancelButton = new Button { Text = "??", Width = 88, Height = 30, DialogResult = DialogResult.Cancel };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);

            Controls.Add(templateListBox);
            Controls.Add(buttons);
            Controls.Add(titleLabel);
            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        public string SelectedTemplateId { get; private set; } = string.Empty;

        private void OkButton_Click(object sender, EventArgs e)
        {
            var selected = templateListBox.SelectedItem as TemplateListItem;
            if (selected == null)
            {
                OperationResultDialog.ShowWarning("????????");
                return;
            }

            SelectedTemplateId = selected.TemplateId;
            DialogResult = DialogResult.OK;
            Close();
        }

        private sealed class TemplateListItem
        {
            public TemplateListItem(TemplateDefinition template)
            {
                TemplateId = template.TemplateId;
                DisplayText = $"{template.TemplateName}  (v{template.Revision})";
            }

            public string TemplateId { get; }
            public string DisplayText { get; }

            public override string ToString() => DisplayText;
        }
    }
}
```

- [ ] **Step 2: Add a simple name-capture dialog for ??????**

```csharp
using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplateNameDialog : Form
    {
        private readonly TextBox templateNameTextBox;

        public TemplateNameDialog(string suggestedTemplateName)
        {
            Text = "????";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(420, 140);

            var promptLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 40,
                Padding = new Padding(12, 12, 12, 0),
                Text = "????????????????????????",
            };

            templateNameTextBox = new TextBox
            {
                Dock = DockStyle.Top,
                Margin = new Padding(12),
                Text = suggestedTemplateName ?? string.Empty,
            };

            var okButton = new Button { Text = "??", Width = 88, Height = 30 };
            okButton.Click += OkButton_Click;
            var cancelButton = new Button { Text = "??", Width = 88, Height = 30, DialogResult = DialogResult.Cancel };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);

            Controls.Add(templateNameTextBox);
            Controls.Add(buttons);
            Controls.Add(promptLabel);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        public string TemplateName { get; private set; } = string.Empty;

        private void OkButton_Click(object sender, EventArgs e)
        {
            var value = (templateNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                OperationResultDialog.ShowWarning("?????????");
                return;
            }

            TemplateName = value;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
```

- [ ] **Step 3: Compose the template store, catalog, and controller in `ThisAddIn` and refresh them on workbook/sheet events**

```csharp
using OfficeAgent.Core.Templates;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn
{
    public partial class ThisAddIn
    {
        internal ITemplateStore TemplateStore { get; private set; }
        internal ITemplateCatalog TemplateCatalog { get; private set; }
        internal RibbonTemplateController RibbonTemplateController { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var appDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAgent");

            TemplateStore = new LocalJsonTemplateStore(Path.Combine(appDataDirectory, "templates"));
            TemplateCatalog = new WorksheetTemplateCatalog(
                SystemConnectorRegistry,
                WorksheetMetadataStore,
                (IWorksheetTemplateBindingStore)WorksheetMetadataStore,
                TemplateStore);
            RibbonTemplateController = new RibbonTemplateController(
                TemplateCatalog,
                GetActiveWorksheetName);

            RibbonTemplateController.RefreshActiveTemplateStateFromSheetMetadata();
            Globals.Ribbons.AgentRibbon?.BindToControllersAndRefresh();
        }

        private void Application_SheetActivate(object sh)
        {
            var sheetName = GetWorksheetName(sh);
            RibbonSyncController?.RefreshProjectFromSheetMetadata(sheetName);
            RibbonTemplateController?.RefreshTemplateState(sheetName);
            lastProjectRefreshSheetName = sheetName;
        }

        private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            RibbonSyncController?.InvalidateRefreshState();
            RibbonTemplateController?.InvalidateRefreshState();
            RibbonSyncController?.RefreshActiveProjectFromSheetMetadata();
            RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
            lastProjectRefreshSheetName = GetActiveWorksheetName();
        }

        private void Application_SheetChange(object sh, Microsoft.Office.Interop.Excel.Range target)
        {
            var sheetName = GetWorksheetName(sh);
            if (!string.Equals(sheetName, "AI_Setting", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var metadataStore = WorksheetMetadataStore as OfficeAgent.ExcelAddIn.Excel.WorksheetMetadataStore;
            metadataStore.InvalidateCache();
            RibbonSyncController?.InvalidateRefreshState();
            RibbonTemplateController?.InvalidateRefreshState();
            lastProjectRefreshSheetName = string.Empty;
        }
    }
}
```

- [ ] **Step 4: Build the add-in test project to make sure the new dialogs and startup composition compile**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~RibbonTemplateControllerTests|FullyQualifiedName~WorksheetMetadataStoreTests"
```

Expected:
- PASS

- [ ] **Step 5: Commit the concrete dialogs and startup wiring**

```bash
git add -- src/OfficeAgent.ExcelAddIn/Dialogs/TemplatePickerDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/TemplateNameDialog.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs
git commit -m "feat: compose ribbon sync template dialogs"
```

### Task 6: Add Ribbon Buttons and Bind Them to the Template Controller

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Extend configuration tests so the ribbon contract is locked before editing designer code**

```csharp
[Fact]
public void TemplateGroupAppearsAfterProjectGroupAndBeforeDownloadGroup()
{
    var designerText = File.ReadAllText(ResolveRepositoryPath(
        "src",
        "OfficeAgent.ExcelAddIn",
        "AgentRibbon.Designer.cs"));

    var projectGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupProject);", StringComparison.Ordinal);
    var templateGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupTemplate);", StringComparison.Ordinal);
    var downloadGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupDownload);", StringComparison.Ordinal);

    Assert.True(projectGroupIndex >= 0);
    Assert.True(templateGroupIndex > projectGroupIndex);
    Assert.True(downloadGroupIndex > templateGroupIndex);
}

[Fact]
public void TemplateGroupContainsApplySaveAndSaveAsButtons()
{
    var designerText = File.ReadAllText(ResolveRepositoryPath(
        "src",
        "OfficeAgent.ExcelAddIn",
        "AgentRibbon.Designer.cs"));

    Assert.Contains("this.groupTemplate.Items.Add(this.applyTemplateButton);", designerText, StringComparison.Ordinal);
    Assert.Contains("this.groupTemplate.Items.Add(this.saveTemplateButton);", designerText, StringComparison.Ordinal);
    Assert.Contains("this.groupTemplate.Items.Add(this.saveAsTemplateButton);", designerText, StringComparison.Ordinal);
}

[Fact]
public void RibbonBindsToTemplateControllerAndRefreshesTemplateButtons()
{
    var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
        "src",
        "OfficeAgent.ExcelAddIn",
        "AgentRibbon.cs"));

    Assert.Contains("TryBindToTemplateController()", ribbonCodeText, StringComparison.Ordinal);
    Assert.Contains("RefreshTemplateButtonsFromController();", ribbonCodeText, StringComparison.Ordinal);
    Assert.Contains("ApplyTemplateButton_Click", ribbonCodeText, StringComparison.Ordinal);
    Assert.Contains("SaveTemplateButton_Click", ribbonCodeText, StringComparison.Ordinal);
    Assert.Contains("SaveAsTemplateButton_Click", ribbonCodeText, StringComparison.Ordinal);
}
```

- [ ] **Step 2: Run the configuration tests to verify they fail**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonConfigurationTests"
```

Expected:
- FAIL because the template ribbon group and controller binding code do not exist yet

- [ ] **Step 3: Add the template ribbon group and hook the buttons into `RibbonTemplateController`**

```csharp
namespace OfficeAgent.ExcelAddIn
{
    partial class AgentRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private void InitializeComponent()
        {
            this.groupTemplate = Factory.CreateRibbonGroup();
            this.applyTemplateButton = Factory.CreateRibbonButton();
            this.saveTemplateButton = Factory.CreateRibbonButton();
            this.saveAsTemplateButton = Factory.CreateRibbonButton();

            this.tab1.Groups.Add(this.groupProject);
            this.tab1.Groups.Add(this.groupTemplate);
            this.tab1.Groups.Add(this.groupDownload);

            this.groupTemplate.Items.Add(this.applyTemplateButton);
            this.groupTemplate.Items.Add(this.saveTemplateButton);
            this.groupTemplate.Items.Add(this.saveAsTemplateButton);
            this.groupTemplate.Label = "??";
            this.groupTemplate.Name = "groupTemplate";

            this.applyTemplateButton.Label = "????";
            this.applyTemplateButton.Name = "applyTemplateButton";
            this.applyTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ApplyTemplateButton_Click);

            this.saveTemplateButton.Label = "????";
            this.saveTemplateButton.Name = "saveTemplateButton";
            this.saveTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveTemplateButton_Click);

            this.saveAsTemplateButton.Label = "????";
            this.saveAsTemplateButton.Name = "saveAsTemplateButton";
            this.saveAsTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsTemplateButton_Click);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton applyTemplateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveTemplateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveAsTemplateButton;
    }
}

namespace OfficeAgent.ExcelAddIn
{
    public partial class AgentRibbon
    {
        private bool isBoundToTemplateController;

        private void AgentRibbon_Load(object sender, Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs e)
        {
            SetProjectDropDownText("?????");
            BindToControllersAndRefresh();
        }

        internal void BindToControllersAndRefresh()
        {
            if (TryBindToSyncController())
            {
                Globals.ThisAddIn.RibbonSyncController?.RefreshActiveProjectFromSheetMetadata();
                RefreshProjectDropDownFromController();
            }

            if (TryBindToTemplateController())
            {
                Globals.ThisAddIn.RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
                RefreshTemplateButtonsFromController();
            }
        }

        private bool TryBindToTemplateController()
        {
            if (isBoundToTemplateController)
            {
                return true;
            }

            var controller = Globals.ThisAddIn.RibbonTemplateController;
            if (controller == null)
            {
                return false;
            }

            controller.TemplateStateChanged += TemplateController_TemplateStateChanged;
            isBoundToTemplateController = true;
            return true;
        }

        internal void RefreshTemplateButtonsFromController()
        {
            var controller = Globals.ThisAddIn.RibbonTemplateController;
            applyTemplateButton.Enabled = controller?.CanApplyTemplate == true;
            saveTemplateButton.Enabled = controller?.CanSaveTemplate == true;
            saveAsTemplateButton.Enabled = controller?.CanSaveAsTemplate == true;
        }

        private void TemplateController_TemplateStateChanged(object sender, EventArgs e)
        {
            RefreshTemplateButtonsFromController();
        }

        private void ApplyTemplateButton_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.ExecuteApplyTemplate();
        }

        private void SaveTemplateButton_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.ExecuteSaveTemplate();
        }

        private void SaveAsTemplateButton_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.ExecuteSaveAsTemplate();
        }

        private void SyncController_ActiveProjectChanged(object sender, EventArgs e)
        {
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (string.IsNullOrWhiteSpace(syncController?.ActiveProjectId))
            {
                ResetProjectDropDownItemsToPlaceholderOnly();
            }
            else
            {
                RebuildProjectDropDownItemsFromCurrentState();
            }

            Globals.ThisAddIn.RibbonTemplateController?.InvalidateRefreshState();
            Globals.ThisAddIn.RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
            RefreshProjectDropDownFromController();
            RefreshTemplateButtonsFromController();
        }
    }
}
```

- [ ] **Step 4: Run the configuration and controller tests to verify the ribbon wiring passes**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonConfigurationTests|FullyQualifiedName~RibbonTemplateControllerTests"
```

Expected:
- PASS

- [ ] **Step 5: Commit the ribbon wiring**

```bash
git add -- src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "feat: add ribbon sync template actions"
```

### Task 7: Document the New Template Workflow and Run Full Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Update the current-behavior snapshot to describe `TemplateBindings` and the template buttons**

```markdown
## 3. ?????

??????????????? `AI_Setting` ??

??????????

- `TemplateBindings`
- `SheetBindings`
- `SheetFieldMappings`

`TemplateBindings` ?????? sheet ?????????????
`SheetBindings + SheetFieldMappings` ?????????????????????

## 4. Ribbon ??

?? Ribbon ?????

- ??
- ??
  - `????`
  - `????`
  - `????`
- ??
- ??
```

- [ ] **Step 2: Update the real-system guide to explain the local-template layer**

```markdown
### 7.6 ????????? metadata ???

?? Ribbon Sync ??????????

- ??????? `%LocalAppData%\OfficeAgent\templates\...`
- `AI_Setting.TemplateBindings` ????? sheet ??????
- ????????????????????? `AI_Setting` ????? `SheetBindings + SheetFieldMappings`

??????????

- ????????????????????
- ????? `AI_Setting` ??????????????
```

- [ ] **Step 3: Extend the manual checklist with apply/save/save-as coverage**

```markdown
## Ribbon Sync ??

- ??????????????????????? `????` ???????????
- ??????? `AI_Setting`?????? `TemplateBindings` section?? `SheetBindings` / `SheetFieldMappings` ????????
- ???? `AI_Setting` ??????????? `????`?????????????????????
- ????????????? `????`????????????????????? sheet ? `TemplateBindings.TemplateId` ???????
- ????? workbook??? `TemplateBindings` section???????????????????
```

- [ ] **Step 4: Run the full automated test matrix**

Run:
```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj
```

Expected:
- PASS across all four test projects

- [ ] **Step 5: Rebuild the add-in and run the manual VSTO flow**

Run:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel
```

Expected:
- Frontend `dist` rebuilt
- Debug VSTO add-in rebuilt
- Excel registration refreshed without build errors

Then execute the new manual checklist items in `docs/vsto-manual-test-checklist.md`.

- [ ] **Step 6: Commit docs and verification updates**

```bash
git add -- docs/modules/ribbon-sync-current-behavior.md docs/ribbon-sync-real-system-integration-guide.md docs/vsto-manual-test-checklist.md
git commit -m "docs: document ribbon sync local templates"
```
