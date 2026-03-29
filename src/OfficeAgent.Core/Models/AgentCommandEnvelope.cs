using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    public static class AgentRouteTypes
    {
        public const string Chat = "chat";
        public const string ExcelCommand = "excelCommand";
        public const string Skill = "skill";
    }

    public static class SkillNames
    {
        public const string UploadData = "upload_data";
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class AgentCommandEnvelope
    {
        public string UserInput { get; set; } = string.Empty;

        public string SkillName { get; set; } = string.Empty;

        public bool Confirmed { get; set; }

        public UploadPreview UploadPreview { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class AgentCommandResult
    {
        public string Route { get; set; } = string.Empty;

        public string SkillName { get; set; } = string.Empty;

        public bool RequiresConfirmation { get; set; }

        public string Status { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public ExcelCommandPreview Preview { get; set; }

        public UploadPreview UploadPreview { get; set; }
    }
}
