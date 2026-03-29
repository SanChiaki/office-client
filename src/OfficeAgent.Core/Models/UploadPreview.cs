using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class UploadPreview
    {
        public string ProjectName { get; set; } = string.Empty;

        public string SheetName { get; set; } = string.Empty;

        public string Address { get; set; } = string.Empty;

        public string[] Headers { get; set; } = System.Array.Empty<string>();

        public string[][] Rows { get; set; } = System.Array.Empty<string[]>();

        public Dictionary<string, string>[] Records { get; set; } = System.Array.Empty<Dictionary<string, string>>();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class UploadExecutionResult
    {
        public int SavedCount { get; set; }

        public string Message { get; set; } = string.Empty;
    }
}
