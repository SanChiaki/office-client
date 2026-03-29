using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IPlanExecutor
    {
        PlanExecutionJournal Execute(AgentPlan plan);
    }
}
