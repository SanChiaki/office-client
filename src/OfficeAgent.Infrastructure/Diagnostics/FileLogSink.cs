using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.Infrastructure.Diagnostics
{
    public sealed class FileLogSink
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        private static readonly JsonSerializerSettings SerializerSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
        };
        private readonly object syncRoot = new object();
        private readonly string logPath;

        public FileLogSink(string logPath)
        {
            this.logPath = logPath ?? throw new ArgumentNullException(nameof(logPath));
        }

        public void Write(OfficeAgentLogEntry entry)
        {
            if (entry == null)
            {
                return;
            }

            var directory = Path.GetDirectoryName(logPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var line = JsonConvert.SerializeObject(entry, SerializerSettings) + Environment.NewLine;
            lock (syncRoot)
            {
                File.AppendAllText(logPath, line, Utf8NoBom);
            }
        }
    }
}
