using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class DevelopmentWorkflowConfigurationTests
    {
        [Fact]
        public void DevRefreshScriptBuildsFrontendAndDebugAddInByDefault()
        {
            var scriptText = File.ReadAllText(ResolveRepositoryPath(
                "eng",
                "Dev-RefreshExcelAddIn.ps1"));

            Assert.Contains("[string]$Configuration = \"Debug\"", scriptText, StringComparison.Ordinal);
            Assert.Contains("Invoke-NativeCommand \"npm.cmd\" \"run\" \"build\"", scriptText, StringComparison.Ordinal);
            Assert.Contains("Build-VstoAddIn.ps1", scriptText, StringComparison.Ordinal);
        }

        [Fact]
        public void DevRefreshScriptCanOptionallyCloseRunningExcelProcesses()
        {
            var scriptText = File.ReadAllText(ResolveRepositoryPath(
                "eng",
                "Dev-RefreshExcelAddIn.ps1"));

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
            Assert.Contains("refresh Excel's local registration", agentsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.ExcelAddIn", agentsText, StringComparison.Ordinal);
        }

        [Fact]
        public void DevRefreshScriptExplainsExcelRegistrationRefresh()
        {
            var scriptText = File.ReadAllText(ResolveRepositoryPath(
                "eng",
                "Dev-RefreshExcelAddIn.ps1"));

            Assert.Contains(
                "Excel registration was refreshed for the development add-in manifest.",
                scriptText,
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
