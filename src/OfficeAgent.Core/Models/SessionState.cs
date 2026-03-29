using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class SessionState
    {
        public string ActiveSessionId { get; set; } = string.Empty;

        public List<ChatSession> Sessions { get; set; } = new List<ChatSession>();
    }
}
