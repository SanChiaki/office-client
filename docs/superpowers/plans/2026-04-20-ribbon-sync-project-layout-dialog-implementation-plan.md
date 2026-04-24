# Ribbon Sync Project Layout Dialog Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a native layout-parameter dialog to Ribbon project selection so first-time binding and project switching require explicit confirmation of `HeaderStartRow`, `HeaderRowCount`, and `DataStartRow` before saving `SheetBindings`.

**Architecture:** Keep the existing Ribbon event flow intact: `AgentRibbon` still forwards the selected `ProjectOption` to `RibbonSyncController`, and `RibbonSyncController` becomes the only place that decides whether the layout dialog is needed. The new WinForms dialog parses and validates raw user input locally, returns a confirmed `SheetBinding` or `null`, and the controller persists only confirmed bindings while restoring the previous active-project state on cancel.

**Tech Stack:** C#, .NET Framework 4.8, WinForms, VSTO Ribbon, xUnit

---

## File Map

- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/ProjectLayoutDialog.cs`
  Purpose: native WinForms dialog for `HeaderStartRow`, `HeaderRowCount`, and `DataStartRow`, plus the input parsing/validation helper used by tests.
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
  Purpose: extend `IRibbonSyncDialogService` and `RibbonSyncDialogService` with the new layout prompt entry point.
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
  Purpose: compile the new dialog source file.
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Purpose: prompt on first bind or project switch, skip same-project reselection, restore the previous dropdown state on cancel, and persist only confirmed bindings.
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs`
  Purpose: lock dialog parsing and validation rules through reflection against the add-in assembly.
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Purpose: lock prompt timing, default-value precedence, cancel rollback, same-project no-op behavior, and ?no auto initialize? behavior.
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  Purpose: keep a regression guard that dropdown rollback still flows through `ActiveProjectChanged` instead of Ribbon-side manual reset logic.
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
  Purpose: describe the new project-selection prompt and cancel/confirm behavior.
- Modify: `docs/vsto-manual-test-checklist.md`
  Purpose: add manual checks for confirm, cancel, validation, and no-auto-initialize behavior.
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`
  Purpose: clarify that connector seed values now prefill the dialog and are saved only after user confirmation.

### Task 1: Lock Dialog Parsing and Validation with Tests

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs`

- [ ] **Step 1: Write failing tests for raw input parsing and layout validation**

Create a focused reflection-based test file so the parser contract is locked before any WinForms code is written:

```csharp
using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ProjectLayoutDialogTests
    {
        [Fact]
        public void TryCreateBindingRejectsNonNumericHeaderStartRow()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "abc", "2", "3", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[4]);
            Assert.Equal("HeaderStartRow ???????", (string)args[5]);
        }

        [Fact]
        public void TryCreateBindingRejectsDataStartInsideHeaderArea()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "2", "2", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[4]);
            Assert.Equal(
                "DataStartRow ??????? HeaderStartRow + HeaderRowCount?",
                (string)args[5]);
        }

        [Fact]
        public void TryCreateBindingReturnsEditedBindingForValidValues()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "4", "1", "5", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.True(success);
            var binding = Assert.IsType<SheetBinding>(args[4]);
            Assert.Equal("Sheet1", binding.SheetName);
            Assert.Equal("current-business-system", binding.SystemKey);
            Assert.Equal("performance", binding.ProjectId);
            Assert.Equal("????", binding.ProjectName);
            Assert.Equal(4, binding.HeaderStartRow);
            Assert.Equal(1, binding.HeaderRowCount);
            Assert.Equal(5, binding.DataStartRow);
            Assert.Null(args[5]);
        }

        private static MethodInfo GetTryCreateBindingMethod()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Dialogs.ProjectLayoutDialog", throwOnError: true)
                .GetMethod(
                    "TryCreateBinding",
                    BindingFlags.Static | BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("ProjectLayoutDialog.TryCreateBinding was not found.");
        }

        private static SheetBinding CreateSeedBinding()
        {
            return new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "????",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
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
    }
}
```

- [ ] **Step 2: Run the targeted test to verify it fails for the intended reason**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ProjectLayoutDialogTests"
```

Expected:
- FAIL with `ProjectLayoutDialog` or `TryCreateBinding` missing

- [ ] **Step 3: Commit the red tests**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs
git commit -m "test: cover ribbon project layout dialog validation"
```

### Task 2: Implement the Dialog and Dialog-Service Plumbing

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/ProjectLayoutDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs`

- [ ] **Step 1: Implement the WinForms dialog and the reusable parser**

Create a code-only dialog so no `.Designer.cs` file is needed, and keep the parsing logic in the same file via a private static helper used by the OK button:

```csharp
using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class ProjectLayoutDialog : Form
    {
        private readonly SheetBinding suggestedBinding;
        private readonly TextBox headerStartRowTextBox;
        private readonly TextBox headerRowCountTextBox;
        private readonly TextBox dataStartRowTextBox;

        public ProjectLayoutDialog(SheetBinding suggestedBinding)
        {
            this.suggestedBinding = suggestedBinding ?? throw new ArgumentNullException(nameof(suggestedBinding));

            Text = "??????";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(420, 250);

            var descriptionLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 60,
                Padding = new Padding(12, 12, 12, 0),
                Text = $"?? sheet?{suggestedBinding.SheetName}\r\n???{FormatProjectLabel(suggestedBinding)}\r\n????????? sheet ??????",
            };

            var layoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 4, 12, 12),
                ColumnCount = 2,
                RowCount = 3,
            };

            headerStartRowTextBox = CreateValueTextBox(suggestedBinding.HeaderStartRow);
            headerRowCountTextBox = CreateValueTextBox(suggestedBinding.HeaderRowCount);
            dataStartRowTextBox = CreateValueTextBox(suggestedBinding.DataStartRow);

            layoutPanel.Controls.Add(new Label { Text = "HeaderStartRow", AutoSize = true }, 0, 0);
            layoutPanel.Controls.Add(headerStartRowTextBox, 1, 0);
            layoutPanel.Controls.Add(new Label { Text = "HeaderRowCount", AutoSize = true }, 0, 1);
            layoutPanel.Controls.Add(headerRowCountTextBox, 1, 1);
            layoutPanel.Controls.Add(new Label { Text = "DataStartRow", AutoSize = true }, 0, 2);
            layoutPanel.Controls.Add(dataStartRowTextBox, 1, 2);

            var okButton = new Button { Text = "??", Width = 88, Height = 30 };
            okButton.Click += OkButton_Click;
            var cancelButton = new Button { Text = "??", Width = 88, Height = 30, DialogResult = DialogResult.Cancel };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };
            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(okButton);

            Controls.Add(layoutPanel);
            Controls.Add(buttonPanel);
            Controls.Add(descriptionLabel);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        public SheetBinding ResultBinding { get; private set; }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (!TryCreateBinding(
                    suggestedBinding,
                    headerStartRowTextBox.Text,
                    headerRowCountTextBox.Text,
                    dataStartRowTextBox.Text,
                    out var binding,
                    out var errorMessage))
            {
                MessageBox.Show(errorMessage, "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ResultBinding = binding;
            DialogResult = DialogResult.OK;
            Close();
        }

        private static bool TryCreateBinding(
            SheetBinding suggestedBinding,
            string headerStartRowText,
            string headerRowCountText,
            string dataStartRowText,
            out SheetBinding binding,
            out string errorMessage)
        {
            binding = null;
            errorMessage = null;

            if (!TryParsePositiveInt(headerStartRowText, "HeaderStartRow", out var headerStartRow, out errorMessage) ||
                !TryParsePositiveInt(headerRowCountText, "HeaderRowCount", out var headerRowCount, out errorMessage) ||
                !TryParsePositiveInt(dataStartRowText, "DataStartRow", out var dataStartRow, out errorMessage))
            {
                return false;
            }

            if (dataStartRow < headerStartRow + headerRowCount)
            {
                errorMessage = "DataStartRow ??????? HeaderStartRow + HeaderRowCount?";
                return false;
            }

            binding = new SheetBinding
            {
                SheetName = suggestedBinding.SheetName,
                SystemKey = suggestedBinding.SystemKey,
                ProjectId = suggestedBinding.ProjectId,
                ProjectName = suggestedBinding.ProjectName,
                HeaderStartRow = headerStartRow,
                HeaderRowCount = headerRowCount,
                DataStartRow = dataStartRow,
            };
            return true;
        }

        private static TextBox CreateValueTextBox(int value)
        {
            return new TextBox
            {
                Width = 120,
                Text = value.ToString(),
            };
        }

        private static string FormatProjectLabel(SheetBinding binding)
        {
            var projectId = binding?.ProjectId ?? string.Empty;
            var projectName = binding?.ProjectName ?? string.Empty;

            if (string.IsNullOrWhiteSpace(projectId))
            {
                return projectName;
            }

            if (string.IsNullOrWhiteSpace(projectName))
            {
                return projectId;
            }

            return projectId + "-" + projectName;
        }

        private static bool TryParsePositiveInt(
            string text,
            string fieldName,
            out int value,
            out string errorMessage)
        {
            value = 0;
            errorMessage = null;

            if (!int.TryParse(text, out value) || value <= 0)
            {
                errorMessage = fieldName + " ???????";
                return false;
            }

            return true;
        }
    }
}
```

- [ ] **Step 2: Extend the dialog service and compile the new source file**

Add the new prompt method to the existing dialog abstraction and wire it through `RibbonSyncDialogService`:

```csharp
using OfficeAgent.Core.Models;

public interface IRibbonSyncDialogService
{
    bool ConfirmDownload(
        string operationName,
        string projectName,
        int rowCount,
        int fieldCount,
        SyncOperationPreview overwritePreview);

    bool ConfirmUpload(string operationName, string projectName, SyncOperationPreview preview);

    SheetBinding ShowProjectLayoutDialog(SheetBinding suggestedBinding);

    void ShowInfo(string message);
    void ShowWarning(string message);
    void ShowError(string message);
}

internal sealed class RibbonSyncDialogService : IRibbonSyncDialogService
{
    public SheetBinding ShowProjectLayoutDialog(SheetBinding suggestedBinding)
    {
        using (var dialog = new ProjectLayoutDialog(suggestedBinding))
        {
            return dialog.ShowDialog() == DialogResult.OK
                ? dialog.ResultBinding
                : null;
        }
    }
}
```

Register the file in the add-in project:

```xml
<Compile Include="Dialogs\DownloadConfirmDialog.cs" />
<Compile Include="Dialogs\OperationResultDialog.cs" />
<Compile Include="Dialogs\ProjectLayoutDialog.cs" />
<Compile Include="Dialogs\UploadConfirmDialog.cs" />
```

- [ ] **Step 3: Run the targeted dialog tests to verify green**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ProjectLayoutDialogTests"
```

Expected:
- PASS

- [ ] **Step 4: Commit the dialog implementation**

```bash
git add src/OfficeAgent.ExcelAddIn/Dialogs/ProjectLayoutDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs
git commit -m "feat: add ribbon project layout dialog"
```

### Task 3: Lock Project-Selection Prompt Semantics with Tests

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Replace the old silent-layout tests with prompt-driven controller tests**

Update `RibbonSyncControllerTests.cs` so the controller contract is explicit:

```csharp
[Fact]
public void SelectProjectShowsLayoutDialogAndSavesConfirmedBindingWithoutAutoInitialize()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var dialogService = new FakeDialogService
    {
        NextProjectLayoutBinding = new SheetBinding
        {
            SheetName = "Sheet1",
            SystemKey = "current-business-system",
            ProjectId = "performance",
            ProjectName = "????",
            HeaderStartRow = 4,
            HeaderRowCount = 1,
            DataStartRow = 5,
        },
    };
    var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");

    InvokeSelectProject(controller, new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "????",
    });

    var prompt = Assert.Single(dialogService.ProjectLayoutPrompts);
    Assert.Equal(1, prompt.HeaderStartRow);
    Assert.Equal(2, prompt.HeaderRowCount);
    Assert.Equal(3, prompt.DataStartRow);
    Assert.Equal(4, metadataStore.LastSavedBinding.HeaderStartRow);
    Assert.Equal(1, metadataStore.LastSavedBinding.HeaderRowCount);
    Assert.Equal(5, metadataStore.LastSavedBinding.DataStartRow);
    Assert.Empty(metadataStore.LastSavedFieldMappings);
    Assert.Null(connector.LastBuildFieldMappingSeedProjectId);
}

[Fact]
public void SelectProjectUsesExistingLayoutAsDialogDefaultsWhenSwitchingProject()
{
    var metadataStore = new FakeWorksheetMetadataStore();
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "old-project",
        ProjectName = "???",
        HeaderStartRow = 5,
        HeaderRowCount = 2,
        DataStartRow = 7,
    };
    var dialogService = new FakeDialogService
    {
        NextProjectLayoutBinding = new SheetBinding
        {
            SheetName = "Sheet1",
            SystemKey = "current-business-system",
            ProjectId = "performance",
            ProjectName = "????",
            HeaderStartRow = 5,
            HeaderRowCount = 2,
            DataStartRow = 7,
        },
    };
    var controller = CreateController(new FakeSystemConnector(), metadataStore, dialogService, () => "Sheet1");

    InvokeSelectProject(controller, new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "????",
    });

    var prompt = Assert.Single(dialogService.ProjectLayoutPrompts);
    Assert.Equal("performance", prompt.ProjectId);
    Assert.Equal("????", prompt.ProjectName);
    Assert.Equal(5, prompt.HeaderStartRow);
    Assert.Equal(2, prompt.HeaderRowCount);
    Assert.Equal(7, prompt.DataStartRow);
}

[Fact]
public void SelectProjectDoesNotPromptOrSaveWhenSameProjectIsReselected()
{
    var metadataStore = new FakeWorksheetMetadataStore();
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "????",
        HeaderStartRow = 5,
        HeaderRowCount = 2,
        DataStartRow = 7,
    };
    var dialogService = new FakeDialogService();
    var controller = CreateController(new FakeSystemConnector(), metadataStore, dialogService, () => "Sheet1");

    InvokeRefresh(controller);
    InvokeSelectProject(controller, new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "????",
    });

    Assert.Empty(dialogService.ProjectLayoutPrompts);
    Assert.Null(metadataStore.LastSavedBinding);
    Assert.Equal("performance", ReadActiveProjectId(controller));
    Assert.Equal("????", ReadActiveProjectDisplayName(controller));
}

[Fact]
public void SelectProjectCancelKeepsExistingBindingAndActiveProjectState()
{
    var metadataStore = new FakeWorksheetMetadataStore();
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "old-project",
        ProjectName = "???",
        HeaderStartRow = 4,
        HeaderRowCount = 1,
        DataStartRow = 5,
    };
    var dialogService = new FakeDialogService
    {
        NextProjectLayoutBinding = null,
    };
    var controller = CreateController(new FakeSystemConnector(), metadataStore, dialogService, () => "Sheet1");

    InvokeRefresh(controller);
    InvokeSelectProject(controller, new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "????",
    });

    Assert.Single(dialogService.ProjectLayoutPrompts);
    Assert.Null(metadataStore.LastSavedBinding);
    Assert.Equal("old-project", ReadActiveProjectId(controller));
    Assert.Equal("???", ReadActiveProjectDisplayName(controller));
}
```

Extend the fake dialog proxy so tests can inspect prompt defaults:

```csharp
public List<SheetBinding> ProjectLayoutPrompts { get; } = new List<SheetBinding>();

public SheetBinding NextProjectLayoutBinding { get; set; }

case "ShowProjectLayoutDialog":
    var suggested = (SheetBinding)call.InArgs[0];
    ProjectLayoutPrompts.Add(CloneBinding(suggested));
    return new ReturnMessage(NextProjectLayoutBinding, null, 0, call.LogicalCallContext, call);

private static SheetBinding CloneBinding(SheetBinding binding)
{
    return binding == null
        ? null
        : new SheetBinding
        {
            SheetName = binding.SheetName,
            SystemKey = binding.SystemKey,
            ProjectId = binding.ProjectId,
            ProjectName = binding.ProjectName,
            HeaderStartRow = binding.HeaderStartRow,
            HeaderRowCount = binding.HeaderRowCount,
            DataStartRow = binding.DataStartRow,
        };
}
```

Add one Ribbon-level regression guard in `AgentRibbonConfigurationTests.cs` so rollback stays controller-driven:

```csharp
[Fact]
public void ProjectSelectionLeavesDropdownResetToControllerRefreshFlow()
{
    var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
        "src",
        "OfficeAgent.ExcelAddIn",
        "AgentRibbon.cs"));

    var methodStart = ribbonCodeText.IndexOf("private void ProjectDropDown_TextChanged", StringComparison.Ordinal);
    var nextMethodStart = ribbonCodeText.IndexOf("private void InitializeSheetButton_Click", methodStart, StringComparison.Ordinal);
    var methodText = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);

    Assert.Contains("Globals.ThisAddIn.RibbonSyncController?.SelectProject(project);", methodText, StringComparison.Ordinal);
    Assert.DoesNotContain("SetProjectDropDownText(", methodText, StringComparison.Ordinal);
    Assert.DoesNotContain("RefreshProjectDropDownFromController();", methodText, StringComparison.Ordinal);
}
```

- [ ] **Step 2: Run the targeted controller/configuration tests to verify they fail for the intended reason**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~RibbonSyncControllerTests|FullyQualifiedName~AgentRibbonConfigurationTests"
```

Expected:
- FAIL because `RibbonSyncController.SelectProject` still saves immediately and never calls `ShowProjectLayoutDialog`

- [ ] **Step 3: Commit the red controller tests**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "test: cover ribbon project selection layout prompt"
```

### Task 4: Implement Prompt-on-Select, Cancel Rollback, and Same-Project No-Op

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Refactor `SelectProject` so it prompts only when a real switch is happening**

Update the controller to load the current binding first, short-circuit same-project reselection, and restore previous state on cancel:

```csharp
public void SelectProject(ProjectOption project)
{
    if (project == null)
    {
        throw new ArgumentNullException(nameof(project));
    }

    var sheetName = GetRequiredSheetName();
    var existingBinding = TryLoadBinding(sheetName);

    if (IsSameProject(existingBinding, project))
    {
        lastRefreshedSheetName = sheetName;
        ApplyBindingState(existingBinding);
        return;
    }

    var suggestedBinding = worksheetSyncService.CreateBindingSeed(sheetName, project);
    var confirmedBinding = dialogService.ShowProjectLayoutDialog(suggestedBinding);

    if (confirmedBinding == null)
    {
        RestoreBindingState(existingBinding, sheetName);
        return;
    }

    metadataStore.SaveBinding(confirmedBinding);
    lastRefreshedSheetName = sheetName;
    ApplyBindingState(confirmedBinding);
}
```

Add the helper methods the tests depend on:

```csharp
private SheetBinding TryLoadBinding(string sheetName)
{
    try
    {
        return metadataStore.LoadBinding(sheetName);
    }
    catch (InvalidOperationException)
    {
        return null;
    }
}

private static bool IsSameProject(SheetBinding existingBinding, ProjectOption project)
{
    return existingBinding != null &&
        string.Equals(existingBinding.SystemKey, project.SystemKey, StringComparison.Ordinal) &&
        string.Equals(existingBinding.ProjectId, project.ProjectId, StringComparison.Ordinal);
}

private void RestoreBindingState(SheetBinding binding, string sheetName)
{
    lastRefreshedSheetName = sheetName;
    if (binding == null)
    {
        ClearActiveProjectState();
        return;
    }

    ApplyBindingState(binding);
}
```

- [ ] **Step 2: Stop injecting placeholder text into bound-but-unnamed projects**

Keep the placeholder only for the ?no binding? state so the existing Ribbon fallback formatter can degrade to plain `ProjectId` when `ProjectName` is blank:

```csharp
private void ApplyBindingState(SheetBinding binding)
{
    if (binding == null)
    {
        ClearActiveProjectState();
        return;
    }

    ActiveProjectId = binding.ProjectId ?? string.Empty;
    ActiveSystemKey = binding.SystemKey ?? string.Empty;
    ActiveProjectDisplayName = binding.ProjectName ?? string.Empty;
    OnActiveProjectChanged();
}

private void ClearActiveProjectState()
{
    ActiveProjectId = string.Empty;
    ActiveSystemKey = string.Empty;
    ActiveProjectDisplayName = DefaultProjectDisplayName;
    OnActiveProjectChanged();
}
```

Delete the old `BuildBindingForSelection` method after `SelectProject` stops using it, because `WorksheetSyncService.CreateBindingSeed(...)` already performs the ?existing layout overrides connector seed? merge.

- [ ] **Step 3: Run the targeted selection-flow tests to verify green**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ProjectLayoutDialogTests|FullyQualifiedName~RibbonSyncControllerTests|FullyQualifiedName~AgentRibbonConfigurationTests"
```

Expected:
- PASS

- [ ] **Step 4: Commit the controller behavior change**

```bash
git add src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "feat: prompt for layout when ribbon project changes"
```

### Task 5: Update User-Facing Docs and Run Final Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/vsto-manual-test-checklist.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`

- [ ] **Step 1: Document the new selection flow and manual verification path**

Update the Ribbon Sync behavior snapshot:

```md
- ?????????? sheet ????????????????????????
- ??????????? sheet ???? `HeaderStartRow`?`HeaderRowCount`?`DataStartRow`???????? `CreateBindingSeed()`
- ???? `??` ???????????????????????
- ???? `??` ?????????? `SheetBindings` ?? `AI_Setting`
- ???????????????????? binding
```

Update the manual checklist:

```md
- ????? sheet ????????????????????? `1 / 2 / 3`
- ? `HeaderStartRow`?`HeaderRowCount`?`DataStartRow` ???????? `??`??? `AI_Setting` ??????????
- ??????????????????????? sheet ????????
- ?????????? `??`????? sheet ????????????????
- ??????? `DataStartRow < HeaderStartRow + HeaderRowCount`???????????
- ??????????????? sheet ?????????
```

Update the integration guide?s seed section:

```md
- `CreateBindingSeed()` ???????????????????????????
- ????????????????? seed???????? `1 / 2 / 3`
- ???????????????????? sheet ?????
- ??????????`SheetBindings` ???? `AI_Setting`
```

- [ ] **Step 2: Run the final automated verification**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ProjectLayoutDialogTests|FullyQualifiedName~RibbonSyncControllerTests|FullyQualifiedName~AgentRibbonConfigurationTests"
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected:
- PASS for the targeted regression set
- PASS for the full `OfficeAgent.ExcelAddIn.Tests` suite

- [ ] **Step 3: Execute the Excel manual checks from the updated checklist**

Use the `Ribbon Sync` section in `docs/vsto-manual-test-checklist.md`, focusing on:
- first bind confirm path
- project switch cancel path
- validation error path
- no-auto-initialize path

- [ ] **Step 4: Commit the docs and verification sweep**

```bash
git add docs/modules/ribbon-sync-current-behavior.md docs/vsto-manual-test-checklist.md docs/ribbon-sync-real-system-integration-guide.md
git commit -m "docs: describe ribbon project layout prompt"
```
