using System;
using System.IO;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Infrastructure.Diagnostics;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class FileLogSinkTests : IDisposable
    {
        private readonly string tempDirectory;

        public FileLogSinkTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.FileLogSink.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void WriteCreatesTheLogDirectoryAndAppendsJsonLines()
        {
            var logPath = Path.Combine(tempDirectory, "logs", "officeagent.log");
            var sink = new FileLogSink(logPath);

            sink.Write(new OfficeAgentLogEntry
            {
                TimestampUtc = new DateTime(2026, 3, 29, 13, 0, 0, DateTimeKind.Utc),
                Level = "info",
                Component = "bridge",
                EventName = "request.received",
                Message = "Received bridge.executeExcelCommand.",
            });

            Assert.True(File.Exists(logPath));
            var lines = File.ReadAllLines(logPath);
            Assert.Single(lines);

            var payload = JObject.Parse(lines[0]);
            Assert.Equal("info", (string)payload["level"]);
            Assert.Equal("bridge", (string)payload["component"]);
            Assert.Equal("request.received", (string)payload["eventName"]);
            Assert.Equal("Received bridge.executeExcelCommand.", (string)payload["message"]);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }
}
