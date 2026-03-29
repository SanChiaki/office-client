using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IExcelCommandExecutor
    {
        ExcelCommandResult Preview(ExcelCommand command);

        ExcelCommandResult Execute(ExcelCommand command);
    }
}
