using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class ChatSession
    {
        public string Id { get; set; } = string.Empty;

        public string Title { get; set; } = string.Empty;

        public DateTime CreatedAtUtc { get; set; }

        public DateTime UpdatedAtUtc { get; set; }

        public List<ChatMessage> Messages { get; set; } = new List<ChatMessage>();
    }
}
