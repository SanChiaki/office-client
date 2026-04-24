# Dev Refresh Excel Add-In Script Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a single development refresh script that rebuilds frontend assets and the Debug VSTO add-in, with an optional Excel shutdown step, and document the workflow in `AGENTS.md`.

**Architecture:** Add a thin orchestration script under `eng/` that reuses the existing `Build-VstoAddIn.ps1` script for add-in compilation instead of duplicating build logic. Cover the behavior with configuration-style xUnit tests that assert the script contract and the developer documentation contract.

**Tech Stack:** PowerShell, xUnit, existing VSTO build script, repository documentation

---

### Task 1: Lock The Script Contract With Failing Tests

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/DevelopmentWorkflowConfigurationTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

- [ ] **Step 1: Write the failing test**

```csharp
[Fact]
public void DevRefreshScriptBuildsFrontendAndDebugAddInByDefault()
{
    var scriptText = File.ReadAllText(ResolveRepositoryPath("eng", "Dev-RefreshExcelAddIn.ps1"));

    Assert.Contains("[string]$Configuration = \"Debug\"", scriptText, StringComparison.Ordinal);
    Assert.Contains("\"npm.cmd\" \"run\" \"build\"", scriptText, StringComparison.Ordinal);
    Assert.Contains("Build-VstoAddIn.ps1", scriptText, StringComparison.Ordinal);
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`
Expected: FAIL because `eng/Dev-RefreshExcelAddIn.ps1` does not exist yet

- [ ] **Step 3: Add more failing coverage for optional Excel shutdown and AGENTS documentation**

```csharp
[Fact]
public void DevRefreshScriptCanOptionallyCloseRunningExcelProcesses()
{
    var scriptText = File.ReadAllText(ResolveRepositoryPath("eng", "Dev-RefreshExcelAddIn.ps1"));

    Assert.Contains("[switch]$CloseExcel", scriptText, StringComparison.Ordinal);
    Assert.Contains("Get-Process EXCEL", scriptText, StringComparison.Ordinal);
    Assert.Contains("Stop-Process -Force", scriptText, StringComparison.Ordinal);
}

[Fact]
public void AgentsDocumentUsesUnifiedDevRefreshScript()
{
    var agentsText = File.ReadAllText(ResolveRepositoryPath("AGENTS.md"));

    Assert.Contains("eng/Dev-RefreshExcelAddIn.ps1", agentsText, StringComparison.Ordinal);
    Assert.Contains("-CloseExcel", agentsText, StringComparison.Ordinal);
}
```

- [ ] **Step 4: Run test suite again to verify failures are still for missing implementation**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`
Expected: FAIL with missing script / missing documentation assertions

### Task 2: Implement The Unified Development Script

**Files:**
- Create: `eng/Dev-RefreshExcelAddIn.ps1`
- Modify: `eng/Build-VstoAddIn.ps1`

- [ ] **Step 1: Create the script shell with dev-oriented defaults**

```powershell
[CmdletBinding()]
param(
    [string]$Configuration = "Debug",
    [switch]$CloseExcel,
    [switch]$SkipFrontend,
    [switch]$SkipAddIn,
    [string]$VisualStudioMSBuildPath
)
```

- [ ] **Step 2: Add repo-root resolution and a native command helper**

```powershell
$repoRoot = Split-Path -Parent $PSScriptRoot
$frontendRoot = Join-Path $repoRoot "src\\OfficeAgent.Frontend"
$addinProject = Join-Path $repoRoot "src\\OfficeAgent.ExcelAddIn\\OfficeAgent.ExcelAddIn.csproj"
$buildVstoAddInScript = Join-Path $repoRoot "eng\\Build-VstoAddIn.ps1"
```

- [ ] **Step 3: Add optional Excel shutdown logic**

```powershell
if ($CloseExcel) {
    $excelProcesses = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        $excelProcesses | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
}
```

- [ ] **Step 4: Build frontend unless explicitly skipped**

```powershell
if (-not $SkipFrontend) {
    Push-Location $frontendRoot
    try {
        Invoke-NativeCommand "npm.cmd" "run" "build"
    }
    finally {
        Pop-Location
    }
}
```

- [ ] **Step 5: Reuse the existing signed VSTO build script for the add-in build**

```powershell
if (-not $SkipAddIn) {
    Invoke-NativeCommand "pwsh" "-NoProfile" "-ExecutionPolicy" "Bypass" "-File" $buildVstoAddInScript "-ProjectPath" $addinProject "-Configuration" $Configuration "-VisualStudioMSBuildPath" $VisualStudioMSBuildPath
}
```

- [ ] **Step 6: Print short operator guidance for next validation steps**

```powershell
Write-Host "Development refresh complete."
Write-Host "- C# / Ribbon changes: reopen Excel before validating."
Write-Host "- Frontend-only changes: reopen the task pane before validating."
```

### Task 3: Document The Workflow In AGENTS.md

**Files:**
- Modify: `AGENTS.md`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/DevelopmentWorkflowConfigurationTests.cs`

- [ ] **Step 1: Add the new script to the build/development command list**

```markdown
- `pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1` for the recommended dev refresh flow (frontend dist + Debug add-in).
```

- [ ] **Step 2: Add a short recommended workflow note**

```markdown
Recommended development flow:
- Frontend-only changes: run `npm run build`
- Add-in / Ribbon / Excel interop changes: run `eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel`
- Installer validation only: run `installer/OfficeAgent.Setup/build.ps1`
```

- [ ] **Step 3: Re-run tests to verify documentation assertions pass**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`
Expected: PASS

### Task 4: Final Verification

**Files:**
- Verify: `eng/Dev-RefreshExcelAddIn.ps1`
- Verify: `AGENTS.md`

- [ ] **Step 1: Run the focused test suite**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`
Expected: PASS with all tests green

- [ ] **Step 2: Execute the new development script in dry real usage**

Run: `pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -SkipFrontend -SkipAddIn`
Expected: exit code 0 and the guidance output without modifying running Excel

- [ ] **Step 3: Commit**

```bash
git add AGENTS.md eng/Dev-RefreshExcelAddIn.ps1 tests/OfficeAgent.ExcelAddIn.Tests/DevelopmentWorkflowConfigurationTests.cs docs/superpowers/plans/2026-04-21-dev-refresh-excel-addin-script.md
git commit -m "build: add unified dev refresh script"
```
