using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class InstallerBundleConfigurationTests
    {
        [Fact]
        public void SetupBundleWixIncludesOfflinePayloadAndCompressedCab()
        {
            var bundleWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs");

            Assert.Contains("vstor_redist.exe", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x86.msi", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x64.msi", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("Compressed=\"yes\"", bundleWxsText, StringComparison.Ordinal);
        }

        [Fact]
        public void SetupBuildScriptReferencesOfflineBundleOutputsAndExtensions()
        {
            var buildScriptText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.Setup",
                "build.ps1");

            Assert.Contains("OfficeAgent.Setup.exe", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.SetupBundle", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Bal.wixext", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Util.wixext", buildScriptText, StringComparison.Ordinal);
        }

        [Fact]
        public void BundlePrerequisiteReadmeIncludesPrerequisiteInstallerFilenames()
        {
            var prereqReadmeText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "prereqs",
                "README.md");

            Assert.Contains("vstor_redist.exe", prereqReadmeText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", prereqReadmeText, StringComparison.Ordinal);
            Assert.Contains("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", prereqReadmeText, StringComparison.Ordinal);
        }

        [Fact]
        public void DirectMsiProductRetainsPrerequisiteBlockMessages()
        {
            var productWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.Setup",
                "Product.wxs");

            Assert.Contains(
                "requires the Microsoft Visual Studio Tools for Office Runtime 4.0 or later",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "Install the VSTO runtime, then run this installer again.",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "requires the Microsoft Edge WebView2 Runtime",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "Install the Evergreen Runtime or your offline enterprise package, then run this installer again.",
                productWxsText,
                StringComparison.Ordinal);
        }

        [Fact]
        public void AgentGuidanceDocumentsReferenceSetupExeAndUpdatedChecklistFlow()
        {
            var agentsText = ReadRepositoryFile("AGENTS.md");
            var checklistText = ReadRepositoryFile("docs", "vsto-manual-test-checklist.md");

            Assert.Contains("OfficeAgent.Setup.exe", agentsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup.exe", checklistText, StringComparison.Ordinal);
            Assert.DoesNotContain(
                "expects WebView2 runtime preinstallation",
                checklistText,
                StringComparison.Ordinal);
        }

        private static string ReadRepositoryFile(params string[] segments)
        {
            var path = ResolveRepositoryPath(segments);
            Assert.True(File.Exists(path), $"Expected repository file to exist: {path}");
            return File.ReadAllText(path);
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
