using System;
using System.Text.RegularExpressions;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Orchestration
{
    public sealed class AgentOrchestrator : IAgentOrchestrator
    {
        private readonly ISkillRegistry skillRegistry;

        public AgentOrchestrator(ISkillRegistry skillRegistry)
        {
            this.skillRegistry = skillRegistry ?? throw new ArgumentNullException(nameof(skillRegistry));
        }

        public AgentCommandResult Execute(AgentCommandEnvelope envelope)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException(nameof(envelope));
            }

            var skillName = ResolveSkillName(envelope);
            if (!string.IsNullOrWhiteSpace(skillName))
            {
                envelope.SkillName = skillName;
                OfficeAgentLog.Info("agent", "route.skill", $"Routing to skill {skillName}.");
                return skillRegistry.Resolve(skillName).Execute(envelope);
            }

            OfficeAgentLog.Info("agent", "route.chat", "Falling back to chat route.");
            return new AgentCommandResult
            {
                Route = AgentRouteTypes.Chat,
                Status = "completed",
                Message = "General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.",
            };
        }

        private static string ResolveSkillName(AgentCommandEnvelope envelope)
        {
            if (!string.IsNullOrWhiteSpace(envelope.SkillName))
            {
                return envelope.SkillName.Trim();
            }

            var input = envelope.UserInput?.Trim() ?? string.Empty;
            if (input.Length == 0)
            {
                return string.Empty;
            }

            if (input.StartsWith("/upload_data", StringComparison.OrdinalIgnoreCase))
            {
                return SkillNames.UploadData;
            }

            if (input.IndexOf("\u4E0A\u4F20\u5230", StringComparison.Ordinal) >= 0)
            {
                return SkillNames.UploadData;
            }

            if (Regex.IsMatch(input, "(?i)\\bupload\\b.+\\bto\\s+.+$"))
            {
                return SkillNames.UploadData;
            }

            return string.Empty;
        }
    }
}
