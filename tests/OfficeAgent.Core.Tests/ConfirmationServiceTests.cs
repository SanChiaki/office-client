using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class ConfirmationServiceTests
    {
        [Fact]
        public void RequiresConfirmationReturnsFalseForReadSelectionTable()
        {
            var service = new ConfirmationService();

            var requiresConfirmation = service.RequiresConfirmation(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.ReadSelectionTable,
            });

            Assert.False(requiresConfirmation);
        }

        [Fact]
        public void RequiresConfirmationReturnsTrueForWorksheetMutations()
        {
            var service = new ConfirmationService();

            var requiresConfirmation = service.RequiresConfirmation(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.DeleteWorksheet,
                SheetName = "Sheet1",
            });

            Assert.True(requiresConfirmation);
        }

        [Fact]
        public void ValidateRejectsWriteRangeCommandsWithoutATargetAddress()
        {
            var service = new ConfirmationService();

            var error = Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.WriteRange,
                Values = new[]
                {
                    new[] { "Name", "Region" },
                    new[] { "Project A", "CN" },
                },
            }));

            Assert.Contains("target address", error.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ValidateRejectsJaggedWriteRangeValues()
        {
            var service = new ConfirmationService();

            var error = Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.WriteRange,
                TargetAddress = "A1:B2",
                Values = new[]
                {
                    new[] { "Name", "Region" },
                    new[] { "Project A" },
                },
            }));

            Assert.Contains("rectangular", error.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ValidateRejectsConflictingSheetNamesForWriteRange()
        {
            var service = new ConfirmationService();

            var error = Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.WriteRange,
                SheetName = "Sheet1",
                TargetAddress = "Sheet2!A1:B2",
                Values = new[]
                {
                    new[] { "Name", "Region" },
                },
            }));

            Assert.Contains("conflicting sheet names", error.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ValidateRejectsQualifiedWriteRangeAddressesWithoutACellReference()
        {
            var service = new ConfirmationService();

            var error = Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.WriteRange,
                TargetAddress = "Sheet1!",
                Values = new[]
                {
                    new[] { "Name", "Region" },
                },
            }));

            Assert.Contains("cell reference", error.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ValidateRejectsWorksheetCommandsWithoutRequiredNames()
        {
            var service = new ConfirmationService();

            Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.AddWorksheet,
            }));
            Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.RenameWorksheet,
                SheetName = "Sheet1",
                NewSheetName = "Sheet1",
            }));
            Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.DeleteWorksheet,
            }));
        }

        [Fact]
        public void ValidateRejectsUnsupportedCommandTypes()
        {
            var service = new ConfirmationService();

            var error = Assert.Throws<System.ArgumentException>(() => service.Validate(new ExcelCommand
            {
                CommandType = "excel.unknown",
            }));

            Assert.Contains("not supported", error.Message, System.StringComparison.OrdinalIgnoreCase);
        }
    }
}
