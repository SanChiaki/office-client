using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WebMessageRouterTests : IDisposable
    {
        private readonly string tempDirectory;

        public WebMessageRouterTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.Router.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void SaveSettingsRejectsMissingPayloadWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\"}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void SaveSettingsRejectsEmptyObjectPayloadWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{}}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void SaveSettingsRejectsPayloadWithoutApiKeyWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"baseUrl\":\"https://api.internal.example\",\"model\":\"gpt-5-mini\"}}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void GetSettingsRejectsEmptyObjectPayload()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getSettings\",\"requestId\":\"req-1\",\"payload\":{}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }

        private static object CreateRouter(FileSessionStore sessionStore, FileSettingsStore settingsStore)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var routerType = addInAssembly.GetType(
                "OfficeAgent.ExcelAddIn.WebBridge.WebMessageRouter",
                throwOnError: true);
            return Activator.CreateInstance(
                routerType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { sessionStore, settingsStore },
                culture: null);
        }

        private static string InvokeRoute(object router, string requestJson)
        {
            var routeMethod = router.GetType().GetMethod(
                "Route",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string)routeMethod.Invoke(router, new object[] { requestJson });
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
