using System;

namespace OfficeAgent.Core.Diagnostics
{
    public sealed class OfficeAgentLogEntry
    {
        public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;

        public string Level { get; set; } = string.Empty;

        public string Component { get; set; } = string.Empty;

        public string EventName { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public string Details { get; set; } = string.Empty;

        public string Exception { get; set; } = string.Empty;
    }
}
