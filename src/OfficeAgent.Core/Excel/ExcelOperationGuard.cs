using System;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Excel
{
    public static class ExcelOperationGuard
    {
        public static void EnsureSelectionSupportsTableRead(SelectionContext selectionContext, bool hasMergedCells)
        {
            if (selectionContext == null || !selectionContext.HasSelection)
            {
                throw new InvalidOperationException(selectionContext?.WarningMessage ?? "No selection available.");
            }

            if (!selectionContext.IsContiguous)
            {
                throw new InvalidOperationException(selectionContext.WarningMessage ?? "Multiple selection areas are not supported yet.");
            }

            if (hasMergedCells)
            {
                throw new InvalidOperationException("Merged cells are not supported for reading selection tables.");
            }
        }

        public static void EnsureWorksheetAllowsMutation(string worksheetName, string actionDescription, bool isProtected)
        {
            if (!isProtected)
            {
                return;
            }

            var displayName = string.IsNullOrWhiteSpace(worksheetName) ? "The target worksheet" : $"Worksheet \"{worksheetName}\"";
            throw new InvalidOperationException($"{displayName} is protected and cannot {actionDescription}.");
        }

        public static void EnsureWorkbookStructureAllowsMutation(string actionDescription, bool isProtected)
        {
            if (isProtected)
            {
                throw new InvalidOperationException($"The active workbook structure is protected and cannot {actionDescription}.");
            }
        }

        public static void EnsureRangeAllowsWrite(string worksheetName, string address, bool hasMergedCells)
        {
            if (!hasMergedCells)
            {
                return;
            }

            throw new InvalidOperationException($"Merged cells are not supported for writing values to {worksheetName}!{address}.");
        }
    }
}
