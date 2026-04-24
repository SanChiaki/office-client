# OfficeAgent Offline Setup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a single-file offline `OfficeAgent.Setup.exe` that detects and installs VSTO Runtime and WebView2 Runtime when needed before installing the existing OfficeAgent MSI.

**Architecture:** Keep the current per-user MSI as the app installer and add a new WiX Burn bundle that embeds both prerequisite installers plus the x86/x64 MSIs. Extend the existing `installer/OfficeAgent.Setup/build.ps1` pipeline so it validates offline prerequisite payloads, ensures required WiX extensions are available, builds both MSIs, then builds the single-file bootstrapper.

**Tech Stack:** WiX Toolset 6, Burn bundle authoring, PowerShell, xUnit configuration tests, Markdown documentation

---

## File Map

- Create: `tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs`
  Purpose: lock the bundle/build/doc contract with text-based xUnit assertions before implementation.
- Create: `installer/OfficeAgent.SetupBundle/Bundle.wxs`
  Purpose: define the offline Burn bundle, prerequisite detection searches, package chain, and MSI architecture routing.
- Create: `installer/OfficeAgent.SetupBundle/BundleLicense.rtf`
  Purpose: provide a simple offline-compatible license/notice payload for the standard WiX bootstrapper UI.
- Create: `installer/OfficeAgent.SetupBundle/README.md`
  Purpose: document how to stage prerequisites, run the build, and what outputs to expect.
- Create: `installer/OfficeAgent.SetupBundle/prereqs/README.md`
  Purpose: document the exact offline prerequisite filenames that must be staged locally.
- Modify: `.gitignore`
  Purpose: keep `prereqs` binaries out of git while allowing the README to stay tracked.
- Modify: `installer/OfficeAgent.Setup/build.ps1`
  Purpose: add WiX extension bootstrap, prerequisite validation, and final `setup.exe` bundle build.
- Modify: `AGENTS.md`
  Purpose: update repository guidance so the recommended installer output is the offline `setup.exe`.
- Modify: `docs/vsto-manual-test-checklist.md`
  Purpose: update the manual checklist to validate the new bootstrapper flow and preserve direct-MSI guard checks.

### Task 1: Lock The Offline Bundle Contract With Failing Tests

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

- [ ] **Step 1: Write the first failing tests for the new bundle source and build script**

```csharp
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class InstallerBundleConfigurationTests
    {
        [Fact]
        public void OfflineBundleSourceEmbedsPrerequisitesAndBothMsiArchitectures()
        {
            var bundleText = File.ReadAllText(ResolveRepositoryPath(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs"));

            Assert.Contains("vstor_redist.exe", bundleText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", bundleText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", bundleText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x86.msi", bundleText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x64.msi", bundleText, StringComparison.Ordinal);
            Assert.Contains("Compressed=\"yes\"", bundleText, StringComparison.Ordinal);
        }

        [Fact]
        public void InstallerBuildScriptProducesOfflineSetupExecutable()
        {
            var scriptText = File.ReadAllText(ResolveRepositoryPath(
                "installer",
                "OfficeAgent.Setup",
                "build.ps1"));

            Assert.Contains("OfficeAgent.Setup.exe", scriptText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.SetupBundle", scriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Bal.wixext", scriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Util.wixext", scriptText, StringComparison.Ordinal);
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
```

- [ ] **Step 2: Add failing coverage for prerequisite documentation and MSI guardrails**

```csharp
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class InstallerBundleConfigurationTests
    {
        [Fact]
        public void OfflineBundleSourceEmbedsPrerequisitesAndBothMsiArchitectures()
        {
            var bundleText = File.ReadAllText(ResolveRepositoryPath(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs"));

            Assert.Contains("vstor_redist.exe", bundleText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", bundleText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", bundleText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x86.msi", bundleText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x64.msi", bundleText, StringComparison.Ordinal);
            Assert.Contains("Compressed=\"yes\"", bundleText, StringComparison.Ordinal);
        }

        [Fact]
        public void InstallerBuildScriptProducesOfflineSetupExecutable()
        {
            var scriptText = File.ReadAllText(ResolveRepositoryPath(
                "installer",
                "OfficeAgent.Setup",
                "build.ps1"));

            Assert.Contains("OfficeAgent.Setup.exe", scriptText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.SetupBundle", scriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Bal.wixext", scriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Util.wixext", scriptText, StringComparison.Ordinal);
        }

        [Fact]
        public void PrerequisiteReadmeDocumentsExactOfflinePayloadNames()
        {
            var prereqReadme = File.ReadAllText(ResolveRepositoryPath(
                "installer",
                "OfficeAgent.SetupBundle",
                "prereqs",
                "README.md"));

            Assert.Contains("vstor_redist.exe", prereqReadme, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", prereqReadme, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", prereqReadme, StringComparison.Ordinal);
        }

        [Fact]
        public void MsiStillBlocksDirectInstallWhenPrerequisitesAreMissing()
        {
            var productText = File.ReadAllText(ResolveRepositoryPath(
                "installer",
                "OfficeAgent.Setup",
                "Product.wxs"));

            Assert.Contains(
                "Install the VSTO runtime, then run this installer again.",
                productText,
                StringComparison.Ordinal);
            Assert.Contains(
                "Install the Evergreen Runtime or your offline enterprise package, then run this installer again.",
                productText,
                StringComparison.Ordinal);
        }

        [Fact]
        public void InstallerDocsPromoteSetupExeAsPrimaryUserInstaller()
        {
            var agentsText = File.ReadAllText(ResolveRepositoryPath("AGENTS.md"));
            var checklistText = File.ReadAllText(ResolveRepositoryPath(
                "docs",
                "vsto-manual-test-checklist.md"));

            Assert.Contains("OfficeAgent.Setup.exe", agentsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup.exe", checklistText, StringComparison.Ordinal);
            Assert.DoesNotContain(
                "current MVP deployment flow expects WebView2 runtime preinstallation",
                checklistText,
                StringComparison.Ordinal);
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
```

- [ ] **Step 3: Run the Excel add-in test project to verify the contract fails for the expected reasons**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected:
- FAIL because `installer/OfficeAgent.SetupBundle/Bundle.wxs` does not exist yet
- FAIL because `installer/OfficeAgent.SetupBundle/prereqs/README.md` does not exist yet
- FAIL because `AGENTS.md` and `docs/vsto-manual-test-checklist.md` do not mention `OfficeAgent.Setup.exe` yet

- [ ] **Step 4: Commit the red test contract**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs
git commit -m "test: lock offline setup bundle contract"
```

### Task 2: Add The Offline Bundle Source Tree

**Files:**
- Create: `installer/OfficeAgent.SetupBundle/Bundle.wxs`
- Create: `installer/OfficeAgent.SetupBundle/BundleLicense.rtf`
- Create: `installer/OfficeAgent.SetupBundle/README.md`
- Create: `installer/OfficeAgent.SetupBundle/prereqs/README.md`
- Modify: `.gitignore`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs`

- [ ] **Step 1: Create the bundle authoring file with prerequisite searches, package chain, and architecture routing**

```xml
<?xml version="1.0" encoding="utf-8"?>
<Wix
  xmlns="http://wixtoolset.org/schemas/v4/wxs"
  xmlns:bal="http://wixtoolset.org/schemas/v4/wxs/bal"
  xmlns:util="http://wixtoolset.org/schemas/v4/wxs/util">
  <Bundle
    Name="OfficeAgent for Excel"
    Manufacturer="OfficeAgent"
    Version="$(var.ProductVersion)"
    UpgradeCode="A35DA388-F1A5-4C24-B3F4-CB8E4F625C3A"
    Compressed="yes"
    DisableModify="yes"
    Condition="ProcessorArchitecture &lt;&gt; 12">
    <BootstrapperApplication>
      <bal:WixStandardBootstrapperApplication
        Theme="standard"
        LicenseFile="BundleLicense.rtf" />
    </BootstrapperApplication>

    <util:RegistrySearch
      Id="VstoRuntimeVersion64Search"
      Variable="VstoRuntimeVersion64"
      Root="HKLM"
      Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
      Value="Version"
      Result="value"
      Bitness="always64" />
    <util:RegistrySearch
      Id="VstoRuntimeVersion32Search"
      Variable="VstoRuntimeVersion32"
      Root="HKLM"
      Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
      Value="Version"
      Result="value"
      Bitness="always32" />
    <util:RegistrySearch
      Id="VstoRuntimeInstall64Search"
      Variable="VstoRuntimeInstall64"
      Root="HKLM"
      Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4"
      Value="Install"
      Result="value"
      Bitness="always64" />
    <util:RegistrySearch
      Id="VstoRuntimeInstall32Search"
      Variable="VstoRuntimeInstall32"
      Root="HKLM"
      Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4"
      Value="Install"
      Result="value"
      Bitness="always32" />
    <util:RegistrySearch
      Id="WebView2RuntimeMachine32Search"
      Variable="WebView2RuntimeMachine32"
      Root="HKLM"
      Key="SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
      Value="pv"
      Result="value"
      Bitness="always32" />
    <util:RegistrySearch
      Id="WebView2RuntimeMachine64Search"
      Variable="WebView2RuntimeMachine64"
      Root="HKLM"
      Key="SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
      Value="pv"
      Result="value"
      Bitness="always64" />
    <util:RegistrySearch
      Id="WebView2RuntimeUserSearch"
      Variable="WebView2RuntimeUser"
      Root="HKCU"
      Key="Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
      Value="pv"
      Result="value" />

    <Chain>
      <ExePackage
        Id="VstoRuntime"
        DisplayName="Microsoft Visual Studio Tools for Office Runtime"
        SourceFile="$(var.PrereqRoot)\vstor_redist.exe"
        InstallArguments="/q /norestart"
        Permanent="yes"
        PerMachine="yes"
        Vital="yes"
        DetectCondition="VstoRuntimeVersion64 OR VstoRuntimeVersion32 OR VstoRuntimeInstall64 = 1 OR VstoRuntimeInstall32 = 1" />

      <ExePackage
        Id="WebView2RuntimeX86"
        DisplayName="Microsoft Edge WebView2 Runtime (x86)"
        SourceFile="$(var.PrereqRoot)\MicrosoftEdgeWebView2RuntimeInstallerX86.exe"
        InstallArguments="/silent /install"
        Permanent="yes"
        Vital="yes"
        InstallCondition="NOT VersionNT64"
        DetectCondition="WebView2RuntimeMachine32 OR WebView2RuntimeMachine64 OR WebView2RuntimeUser" />

      <ExePackage
        Id="WebView2RuntimeX64"
        DisplayName="Microsoft Edge WebView2 Runtime (x64)"
        SourceFile="$(var.PrereqRoot)\MicrosoftEdgeWebView2RuntimeInstallerX64.exe"
        InstallArguments="/silent /install"
        Permanent="yes"
        Vital="yes"
        InstallCondition="VersionNT64"
        DetectCondition="WebView2RuntimeMachine32 OR WebView2RuntimeMachine64 OR WebView2RuntimeUser" />

      <MsiPackage
        Id="OfficeAgentMsiX86"
        DisplayName="OfficeAgent for Excel (x86)"
        SourceFile="$(var.MsiRoot)\OfficeAgent.Setup-x86.msi"
        InstallCondition="NOT VersionNT64"
        Visible="no"
        Vital="yes" />

      <MsiPackage
        Id="OfficeAgentMsiX64"
        DisplayName="OfficeAgent for Excel (x64)"
        SourceFile="$(var.MsiRoot)\OfficeAgent.Setup-x64.msi"
        InstallCondition="VersionNT64"
        Visible="no"
        Vital="yes" />
    </Chain>
  </Bundle>
</Wix>
```

- [ ] **Step 2: Add the bundle notice file and both README files**

```rtf
{\rtf1\ansi\deff0
{\fonttbl{\f0 Segoe UI;}}
\f0\fs20 OfficeAgent offline setup notice.\par
This installer may install Microsoft Visual Studio Tools for Office Runtime and Microsoft Edge WebView2 Runtime if they are missing.\par
Continue only if your organization permits installation of these prerequisites.\par
}
```

````markdown
# OfficeAgent Offline Setup Bundle

This directory contains the WiX Burn authoring that produces `artifacts/installer/OfficeAgent.Setup.exe`.

## Required staged prerequisites

Place the following offline installers in `installer/OfficeAgent.SetupBundle/prereqs/` before running the build:

- `vstor_redist.exe`
- `MicrosoftEdgeWebView2RuntimeInstallerX86.exe`
- `MicrosoftEdgeWebView2RuntimeInstallerX64.exe`

## Build command

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1
```

Expected outputs:

- `artifacts/installer/OfficeAgent.Setup.exe`
- `artifacts/installer/OfficeAgent.Setup-x86.msi`
- `artifacts/installer/OfficeAgent.Setup-x64.msi`
````

````markdown
# Offline Prerequisite Payloads

Do not commit the prerequisite installers in this folder.

Before building the offline setup bundle, copy in these exact filenames:

- `vstor_redist.exe`
- `MicrosoftEdgeWebView2RuntimeInstallerX86.exe`
- `MicrosoftEdgeWebView2RuntimeInstallerX64.exe`

The build script fails fast when any of these files are missing.
````

- [ ] **Step 3: Ignore staged prerequisite binaries while keeping the README tracked**

```gitignore
/installer/OfficeAgent.SetupBundle/prereqs/*
!/installer/OfficeAgent.SetupBundle/prereqs/README.md
```

- [ ] **Step 4: Run the targeted tests to verify the bundle source tree satisfies the new contract**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected:
- PASS for `OfflineBundleSourceEmbedsPrerequisitesAndBothMsiArchitectures`
- PASS for `PrerequisiteReadmeDocumentsExactOfflinePayloadNames`
- remaining docs/build-script tests still FAIL because `build.ps1`, `AGENTS.md`, and `docs/vsto-manual-test-checklist.md` are not updated yet

- [ ] **Step 5: Commit the bundle source tree**

```bash
git add .gitignore installer/OfficeAgent.SetupBundle/Bundle.wxs installer/OfficeAgent.SetupBundle/BundleLicense.rtf installer/OfficeAgent.SetupBundle/README.md installer/OfficeAgent.SetupBundle/prereqs/README.md
git commit -m "build: add offline setup bundle authoring"
```

### Task 3: Extend The Installer Build Pipeline To Produce `setup.exe`

**Files:**
- Modify: `installer/OfficeAgent.Setup/build.ps1`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs`

- [ ] **Step 1: Add bundle paths plus reusable helper functions for WiX extensions and required files**

```powershell
$bundleRoot = Join-Path $repoRoot "installer\\OfficeAgent.SetupBundle"
$bundleSource = Join-Path $bundleRoot "Bundle.wxs"
$bundlePrereqRoot = Join-Path $bundleRoot "prereqs"
$offlineSetupPath = Join-Path $outputRoot "OfficeAgent.Setup.exe"
$offlineSetupWixPdbPath = Join-Path $outputRoot "OfficeAgent.Setup.wixpdb"
$toolsManifestPath = Join-Path $repoRoot ".config\\dotnet-tools.json"
$toolsManifest = Get-Content -Raw $toolsManifestPath | ConvertFrom-Json
$wixToolVersion = $toolsManifest.tools.wix.version

function Ensure-WixExtension {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExtensionReference
    )

    Write-Host "Ensuring WiX extension $ExtensionReference..."
    Invoke-NativeCommand "dotnet" "wix" "extension" "add" $ExtensionReference
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (!(Test-Path $Path -PathType Leaf)) {
        throw "Missing $Description: $Path"
    }
}
```

- [ ] **Step 2: Ensure the bundle prerequisites are available before building the bundle**

```powershell
Write-Host "Ensuring WiX bundle extensions are installed..."
Ensure-WixExtension "WixToolset.Bal.wixext/$wixToolVersion"
Ensure-WixExtension "WixToolset.Util.wixext/$wixToolVersion"

$vstoRuntimeInstaller = Join-Path $bundlePrereqRoot "vstor_redist.exe"
$webView2RuntimeInstallerX86 = Join-Path $bundlePrereqRoot "MicrosoftEdgeWebView2RuntimeInstallerX86.exe"
$webView2RuntimeInstallerX64 = Join-Path $bundlePrereqRoot "MicrosoftEdgeWebView2RuntimeInstallerX64.exe"

Assert-FileExists -Path $bundleSource -Description "offline setup bundle source"
Assert-FileExists -Path $vstoRuntimeInstaller -Description "VSTO runtime redistributable"
Assert-FileExists -Path $webView2RuntimeInstallerX86 -Description "WebView2 x86 standalone installer"
Assert-FileExists -Path $webView2RuntimeInstallerX64 -Description "WebView2 x64 standalone installer"
```

- [ ] **Step 3: Build the bundle after the x86/x64 MSIs are produced**

```powershell
Assert-FileExists -Path (Join-Path $outputRoot "OfficeAgent.Setup-x86.msi") -Description "x86 OfficeAgent MSI"
Assert-FileExists -Path (Join-Path $outputRoot "OfficeAgent.Setup-x64.msi") -Description "x64 OfficeAgent MSI"

if (Test-Path $offlineSetupPath) {
    Remove-Item -Force $offlineSetupPath
}

if (Test-Path $offlineSetupWixPdbPath) {
    Remove-Item -Force $offlineSetupWixPdbPath
}

Write-Host "Building offline setup bundle..."
Invoke-NativeCommand "dotnet" "wix" "build" $bundleSource `
    "-arch" "x86" `
    "-d" "ProductVersion=$productVersion" `
    "-d" "PrereqRoot=$bundlePrereqRoot" `
    "-d" "MsiRoot=$outputRoot" `
    "-o" $offlineSetupPath

Write-Host "Installer outputs created at:"
$builtMsiPaths | ForEach-Object { Write-Host " - $_" }
Write-Host " - $offlineSetupPath"
```

- [ ] **Step 4: Run the test suite again to verify the build contract turns green**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected:
- PASS for `InstallerBuildScriptProducesOfflineSetupExecutable`
- docs-related tests may still FAIL until `AGENTS.md` and the manual checklist are updated

- [ ] **Step 5: Commit the build pipeline changes**

```bash
git add installer/OfficeAgent.Setup/build.ps1
git commit -m "build: emit offline setup executable"
```

### Task 4: Update Repository And Manual Validation Documentation

**Files:**
- Modify: `AGENTS.md`
- Modify: `docs/vsto-manual-test-checklist.md`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs`

- [ ] **Step 1: Update `AGENTS.md` so the repository guidance mentions the bundle directory and the new installer output**

```markdown
## Project Structure & Module Organization
`src/OfficeAgent.Core` contains orchestration, models, skills, and service contracts. `src/OfficeAgent.Infrastructure` holds HTTP clients, storage, diagnostics, and DPAPI helpers. `src/OfficeAgent.ExcelAddIn` hosts the ribbon, task pane, Excel interop, and WebView bridge. `src/OfficeAgent.Frontend` is the React/Vite UI. Tests live in `tests/OfficeAgent.Core.Tests`, `tests/OfficeAgent.Infrastructure.Tests`, `tests/OfficeAgent.ExcelAddIn.Tests`, and `tests/OfficeAgent.IntegrationTests`; `tests/mock-server` provides local SSO and API fixtures. Installer sources live in `installer/OfficeAgent.Setup` and `installer/OfficeAgent.SetupBundle`.
```

```markdown
- `pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1` for frontend + add-in + MSI + offline `setup.exe` builds.
```

```markdown
- Installer validation only: run `pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1`, then validate `artifacts/installer/OfficeAgent.Setup.exe`.
```

- [ ] **Step 2: Rewrite the installer section of the manual checklist around the offline bootstrapper**

```markdown
## Installer

- Run `installer/OfficeAgent.Setup/build.ps1` and confirm `artifacts/installer/OfficeAgent.Setup.exe`, `artifacts/installer/OfficeAgent.Setup-x86.msi`, and `artifacts/installer/OfficeAgent.Setup-x64.msi` are created.
- Confirm `installer/OfficeAgent.SetupBundle/prereqs/` contains `vstor_redist.exe`, `MicrosoftEdgeWebView2RuntimeInstallerX86.exe`, and `MicrosoftEdgeWebView2RuntimeInstallerX64.exe` before the build starts.
- Run `OfficeAgent.Setup.exe` on a machine missing both prerequisites and confirm it installs VSTO Runtime, installs WebView2 Runtime, then installs OfficeAgent.
- Run `OfficeAgent.Setup.exe` on a machine with both prerequisites already installed and confirm it skips both prerequisite installers.
- Run `OfficeAgent.Setup.exe` twice on the same machine and confirm the second run does not reinstall the prerequisites and falls through to normal OfficeAgent maintenance behavior.
- Choose the MSI that matches the target Excel bitness only for direct enterprise distribution or debugging scenarios.
- On a machine missing the VSTO runtime, confirm the direct MSI still blocks with a clear prerequisite message.
- On a machine missing the WebView2 runtime, confirm the direct MSI still blocks with a clear prerequisite message.
- Confirm files are deployed under `%LocalAppData%\\OfficeAgent\\ExcelAddIn`.
- Confirm Excel add-in registry entries exist under `HKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\OfficeAgent.ExcelAddIn`.
```

- [ ] **Step 3: Run the test project one more time to verify the documentation contract is green**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected:
- PASS for `InstallerDocsPromoteSetupExeAsPrimaryUserInstaller`
- PASS for the previously-added bundle and build-script configuration tests

- [ ] **Step 4: Commit the documentation updates**

```bash
git add AGENTS.md docs/vsto-manual-test-checklist.md
git commit -m "docs: document offline setup bundle flow"
```

### Task 5: Build And Validate The Offline Installer End-To-End

**Files:**
- Modify: `installer/OfficeAgent.SetupBundle/prereqs/` (local staged binaries only, not committed)
- Test: `installer/OfficeAgent.Setup/build.ps1`
- Test: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Stage the three prerequisite binaries with the exact filenames expected by the bundle**

```powershell
Copy-Item C:\Installers\vstor_redist.exe installer\OfficeAgent.SetupBundle\prereqs\
Copy-Item C:\Installers\MicrosoftEdgeWebView2RuntimeInstallerX86.exe installer\OfficeAgent.SetupBundle\prereqs\
Copy-Item C:\Installers\MicrosoftEdgeWebView2RuntimeInstallerX64.exe installer\OfficeAgent.SetupBundle\prereqs\
```

- [ ] **Step 2: Run the installer build and verify it produces the offline bootstrapper**

Run: `pwsh -NoProfile -ExecutionPolicy Bypass -File installer/OfficeAgent.Setup/build.ps1`

Expected:
- PASS
- output contains `OfficeAgent.Setup-x86.msi`
- output contains `OfficeAgent.Setup-x64.msi`
- output contains `OfficeAgent.Setup.exe`

- [ ] **Step 3: Smoke-check that the expected installer outputs exist**

Run:

```powershell
@(
  'artifacts\installer\OfficeAgent.Setup-x86.msi',
  'artifacts\installer\OfficeAgent.Setup-x64.msi',
  'artifacts\installer\OfficeAgent.Setup.exe'
) | ForEach-Object { "{0} => {1}" -f $_, (Test-Path $_) }
```

Expected:

```text
artifacts\installer\OfficeAgent.Setup-x86.msi => True
artifacts\installer\OfficeAgent.Setup-x64.msi => True
artifacts\installer\OfficeAgent.Setup.exe => True
```

- [ ] **Step 4: Run the installer checklist locally and in an isolated clean environment**

Run locally:
- direct-MSI guard checks from `docs/vsto-manual-test-checklist.md`
- repeat-run `OfficeAgent.Setup.exe` check on the current machine

Run in Windows Sandbox or a disposable VM:
- missing-both-prerequisites install path
- missing-only-VSTO path
- missing-only-WebView2 path
- old-OfficeAgent upgrade path

Expected:
- `setup.exe` installs missing prerequisites only when required
- direct MSI still blocks when a prerequisite is missing
- OfficeAgent loads after installation on both x86 and x64 Excel targets

- [ ] **Step 5: Commit the verified implementation**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/InstallerBundleConfigurationTests.cs .gitignore installer/OfficeAgent.Setup/build.ps1 installer/OfficeAgent.SetupBundle/Bundle.wxs installer/OfficeAgent.SetupBundle/BundleLicense.rtf installer/OfficeAgent.SetupBundle/README.md installer/OfficeAgent.SetupBundle/prereqs/README.md AGENTS.md docs/vsto-manual-test-checklist.md
git commit -m "build: add offline setup bootstrapper"
```
