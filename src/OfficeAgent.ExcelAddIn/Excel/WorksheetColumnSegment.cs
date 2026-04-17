using System;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetColumnSegment
    {
        public int StartColumn { get; set; }

        public int EndColumn { get; set; }

        public WorksheetRuntimeColumn[] Columns { get; set; } = Array.Empty<WorksheetRuntimeColumn>();
    }
}
