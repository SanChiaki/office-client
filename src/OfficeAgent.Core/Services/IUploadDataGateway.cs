using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IUploadDataGateway
    {
        UploadExecutionResult Upload(UploadPreview preview);
    }
}
