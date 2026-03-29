using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
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

        [Fact]
        public void GetSelectionContextReturnsCurrentSelectionContext()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var selectionContextService = new FakeExcelContextService(new SelectionContext
            {
                WorkbookName = "Quarterly Report.xlsx",
                SheetName = "Sheet1",
                Address = "A1:C4",
                RowCount = 4,
                ColumnCount = 3,
                IsContiguous = true,
                HeaderPreview = new[] { "Name", "Region", "Amount" },
                SampleRows = new[]
                {
                    new[] { "Project A", "CN", "42" },
                    new[] { "Project B", "US", "36" },
                },
                WarningMessage = null,
            });

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                selectionContextService,
                new FakeExcelCommandExecutor());
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getSelectionContext\",\"requestId\":\"req-1\"}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"workbookName\":\"Quarterly Report.xlsx\"", responseJson);
            Assert.Contains("\"sheetName\":\"Sheet1\"", responseJson);
            Assert.Contains("\"address\":\"A1:C4\"", responseJson);
        }

        [Fact]
        public void ExecuteExcelCommandExecutesReadCommandsImmediately()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor
            {
                ExecuteResult = new ExcelCommandResult
                {
                    CommandType = ExcelCommandTypes.ReadSelectionTable,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = "Read selection from Sheet1 A1:C4.",
                },
            };

            var router = CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")), executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.readSelectionTable\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"requiresConfirmation\":false", responseJson);
            Assert.Equal(1, executor.ExecuteCalls);
            Assert.Equal(0, executor.PreviewCalls);
        }

        [Fact]
        public void ExecuteExcelCommandReturnsPreviewForUnconfirmedWriteCommands()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor
            {
                PreviewResult = new ExcelCommandResult
                {
                    CommandType = ExcelCommandTypes.AddWorksheet,
                    RequiresConfirmation = true,
                    Status = "preview",
                    Message = "Confirm worksheet creation before Excel is modified.",
                    Preview = new ExcelCommandPreview
                    {
                        Title = "Confirm Excel action",
                        Summary = "Add worksheet \"Summary\"",
                    },
                },
            };

            var router = CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")), executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.addWorksheet\",\"newSheetName\":\"Summary\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"requiresConfirmation\":true", responseJson);
            Assert.Contains("\"summary\":\"Add worksheet", responseJson);
            Assert.Equal(0, executor.ExecuteCalls);
            Assert.Equal(1, executor.PreviewCalls);
        }

        [Fact]
        public void ExecuteExcelCommandRejectsConflictingWriteRangeSheetNames()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor();

            var router = CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")), executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.writeRange\",\"sheetName\":\"Sheet1\",\"targetAddress\":\"Sheet2!A1:B2\",\"values\":[[\"Name\",\"Region\"]],\"confirmed\":false}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"invalid_command\"", responseJson);
            Assert.Equal(0, executor.ExecuteCalls);
            Assert.Equal(0, executor.PreviewCalls);
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
            return CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                new FakeExcelCommandExecutor());
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService selectionContextService,
            IExcelCommandExecutor excelCommandExecutor)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var routerType = addInAssembly.GetType(
                "OfficeAgent.ExcelAddIn.WebBridge.WebMessageRouter",
                throwOnError: true);
            return Activator.CreateInstance(
                routerType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { sessionStore, settingsStore, selectionContextService, excelCommandExecutor },
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

        private sealed class FakeExcelContextService : IExcelContextService
        {
            private readonly SelectionContext selectionContext;

            public FakeExcelContextService(SelectionContext selectionContext)
            {
                this.selectionContext = selectionContext;
            }

            public SelectionContext GetCurrentSelectionContext()
            {
                return selectionContext;
            }
        }

        private sealed class FakeExcelCommandExecutor : IExcelCommandExecutor
        {
            public int ExecuteCalls { get; private set; }

            public int PreviewCalls { get; private set; }

            public ExcelCommandResult ExecuteResult { get; set; } = new ExcelCommandResult
            {
                CommandType = ExcelCommandTypes.ReadSelectionTable,
                RequiresConfirmation = false,
                Status = "completed",
                Message = "Executed.",
            };

            public ExcelCommandResult PreviewResult { get; set; } = new ExcelCommandResult
            {
                CommandType = ExcelCommandTypes.AddWorksheet,
                RequiresConfirmation = true,
                Status = "preview",
                Message = "Preview ready.",
            };

            public ExcelCommandResult Preview(ExcelCommand command)
            {
                PreviewCalls++;
                return PreviewResult;
            }

            public ExcelCommandResult Execute(ExcelCommand command)
            {
                ExecuteCalls++;
                return ExecuteResult;
            }
        }
    }
}
