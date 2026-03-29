using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ILlmPlannerClient
    {
        string Complete(PlannerRequest request);
    }
}
