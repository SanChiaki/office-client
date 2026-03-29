using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ISkillRegistry
    {
        IAgentSkill Resolve(string skillName);
    }

    public interface IAgentSkill
    {
        string SkillName { get; }

        AgentCommandResult Execute(AgentCommandEnvelope envelope);
    }
}
