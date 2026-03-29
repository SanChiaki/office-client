using System;
using OfficeAgent.Core.Excel;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class ExcelOperationGuardTests
    {
        [Fact]
        public void EnsureSelectionSupportsTableReadThrowsForMergedCells()
        {
            var selection = new SelectionContext
            {
                HasSelection = true,
                IsContiguous = true,
                Address = "A1:C3",
            };

            var error = Assert.Throws<InvalidOperationException>(
                () => ExcelOperationGuard.EnsureSelectionSupportsTableRead(selection, hasMergedCells: true));

            Assert.Equal("Merged cells are not supported for reading selection tables.", error.Message);
        }

        [Fact]
        public void EnsureWorksheetAllowsMutationThrowsForProtectedWorksheets()
        {
            var error = Assert.Throws<InvalidOperationException>(
                () => ExcelOperationGuard.EnsureWorksheetAllowsMutation("Summary", "rename worksheets", isProtected: true));

            Assert.Equal("Worksheet \"Summary\" is protected and cannot rename worksheets.", error.Message);
        }

        [Fact]
        public void EnsureSelectionSupportsTableReadAllowsContiguousUnmergedSelections()
        {
            var selection = new SelectionContext
            {
                HasSelection = true,
                IsContiguous = true,
                Address = "A1:C3",
            };

            ExcelOperationGuard.EnsureSelectionSupportsTableRead(selection, hasMergedCells: false);
        }
    }
}
