[CmdletBinding()]
param(
    [string]$Configuration = "Debug",

    [switch]$CloseExcel,

    [switch]$SkipFrontend,

    [switch]$SkipAddIn,

    [string]$VisualStudioMSBuildPath
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$frontendRoot = Join-Path $repoRoot "src\OfficeAgent.Frontend"
$addinProject = Join-Path $repoRoot "src\OfficeAgent.ExcelAddIn\OfficeAgent.ExcelAddIn.csproj"
$buildVstoAddInScript = Join-Path $repoRoot "eng\Build-VstoAddIn.ps1"

function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

if ($CloseExcel) {
    $excelProcesses = Get-Process EXCEL -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        Write-Host "Closing running Excel processes..."
        $excelProcesses | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
    else {
        Write-Host "No running Excel processes were found."
    }
}

if (-not $SkipFrontend) {
    Write-Host "Building frontend dist..."
    Push-Location $frontendRoot
    try {
        Invoke-NativeCommand "npm.cmd" "run" "build"
    }
    finally {
        Pop-Location
    }
}
else {
    Write-Host "Skipping frontend build."
}

if (-not $SkipAddIn) {
    Write-Host "Building Debug VSTO add-in..."
    $buildArgs = @(
        "-NoProfile"
        "-ExecutionPolicy"
        "Bypass"
        "-File"
        $buildVstoAddInScript
        "-ProjectPath"
        $addinProject
        "-Configuration"
        $Configuration
    )

    if (-not [string]::IsNullOrWhiteSpace($VisualStudioMSBuildPath)) {
        $buildArgs += @(
            "-VisualStudioMSBuildPath"
            $VisualStudioMSBuildPath
        )
    }

    Invoke-NativeCommand "pwsh" @buildArgs
}
else {
    Write-Host "Skipping VSTO add-in build."
}

Write-Host "Development refresh complete."
if (-not $SkipAddIn) {
    Write-Host "- Excel registration was refreshed for the development add-in manifest."
    Write-Host "- C# / Ribbon / Excel interop changes: reopen Excel before validating."
}

if (-not $SkipFrontend) {
    Write-Host "- Frontend-only or mixed UI changes: reopen the task pane before validating."
}
