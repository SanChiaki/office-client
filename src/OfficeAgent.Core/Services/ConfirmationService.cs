using System;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public sealed class ConfirmationService
    {
        public bool RequiresConfirmation(ExcelCommand command)
        {
            var commandType = NormalizeCommandType(command);
            return commandType != ExcelCommandTypes.ReadSelectionTable;
        }

        public void Validate(ExcelCommand command)
        {
            var commandType = NormalizeCommandType(command);

            switch (commandType)
            {
                case ExcelCommandTypes.ReadSelectionTable:
                    return;
                case ExcelCommandTypes.WriteRange:
                    ValidateWriteRange(command);
                    return;
                case ExcelCommandTypes.AddWorksheet:
                    RequireValue(command.NewSheetName, "new sheet name");
                    return;
                case ExcelCommandTypes.RenameWorksheet:
                    RequireValue(command.SheetName, "sheet name");
                    RequireValue(command.NewSheetName, "new sheet name");
                    if (string.Equals(command.SheetName, command.NewSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new ArgumentException("Rename worksheet commands must change the worksheet name.");
                    }

                    return;
                case ExcelCommandTypes.DeleteWorksheet:
                    RequireValue(command.SheetName, "sheet name");
                    return;
                default:
                    throw new ArgumentException($"Excel command type '{command.CommandType}' is not supported.");
            }
        }

        private static string NormalizeCommandType(ExcelCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            if (string.IsNullOrWhiteSpace(command.CommandType))
            {
                throw new ArgumentException("Excel commands must include a command type.");
            }

            return command.CommandType.Trim();
        }

        private static void ValidateWriteRange(ExcelCommand command)
        {
            RequireValue(command.TargetAddress, "target address");
            EnsureTargetAddressContainsCellReference(command.TargetAddress);
            ValidateSheetNameQualifier(command);

            if (command.Values == null || command.Values.Length == 0)
            {
                throw new ArgumentException("Write range commands require at least one row of values.");
            }

            var expectedColumnCount = command.Values[0]?.Length ?? 0;
            if (expectedColumnCount == 0)
            {
                throw new ArgumentException("Write range commands require at least one column of values.");
            }

            if (command.Values.Any((row) => row == null || row.Length != expectedColumnCount))
            {
                throw new ArgumentException("Write range commands require a rectangular values payload.");
            }
        }

        private static void EnsureTargetAddressContainsCellReference(string targetAddress)
        {
            var address = targetAddress.Trim();
            if (!address.Contains("!"))
            {
                return;
            }

            var segments = address.Split(new[] { '!' }, 2);
            if (segments.Length == 2 && string.IsNullOrWhiteSpace(segments[1]))
            {
                throw new ArgumentException("Write range commands must include a cell reference in the target address.");
            }
        }

        private static void RequireValue(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException($"Excel commands require a {fieldName}.");
            }
        }

        private static void ValidateSheetNameQualifier(ExcelCommand command)
        {
            if (string.IsNullOrWhiteSpace(command.SheetName) || string.IsNullOrWhiteSpace(command.TargetAddress))
            {
                return;
            }

            var address = command.TargetAddress.Trim();
            if (!address.Contains("!"))
            {
                return;
            }

            var segments = address.Split(new[] { '!' }, 2);
            if (segments.Length != 2 || string.IsNullOrWhiteSpace(segments[0]))
            {
                return;
            }

            if (!string.Equals(command.SheetName.Trim(), segments[0].Trim(), StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Write range commands cannot specify conflicting sheet names.");
            }
        }
    }
}
