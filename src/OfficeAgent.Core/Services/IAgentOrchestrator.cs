using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAgentOrchestrator
    {
        AgentCommandResult Execute(AgentCommandEnvelope envelope);
    }
}
