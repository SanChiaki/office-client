using System;

namespace OfficeAgent.Core.Models
{
    public sealed class ChatMessage
    {
        public string Id { get; set; } = string.Empty;

        public string Role { get; set; } = string.Empty;

        public string Content { get; set; } = string.Empty;

        public DateTime CreatedAtUtc { get; set; }
    }
}
