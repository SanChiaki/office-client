using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetHeaderMatcher
    {
        private readonly FieldMappingValueAccessor valueAccessor;

        public WorksheetHeaderMatcher(FieldMappingValueAccessor valueAccessor)
        {
            this.valueAccessor = valueAccessor ?? throw new ArgumentNullException(nameof(valueAccessor));
        }

        public WorksheetRuntimeColumn[] Match(
            string sheetName,
            SheetBinding binding,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings,
            IWorksheetGridAdapter grid)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            if (grid == null)
            {
                throw new ArgumentNullException(nameof(grid));
            }

            var rows = mappings ?? Array.Empty<SheetFieldMappingRow>();
            var lookup = BuildLookup(definition, rows);
            var result = new List<WorksheetRuntimeColumn>();
            var lastUsedColumn = grid.GetLastUsedColumn(sheetName);
            var headerRow = binding.HeaderStartRow <= 0 ? 1 : binding.HeaderStartRow;
            var currentParent = string.Empty;

            for (var column = 1; column <= lastUsedColumn; column++)
            {
                var topText = grid.GetCellText(sheetName, headerRow, column) ?? string.Empty;
                var bottomText = binding.HeaderRowCount > 1
                    ? grid.GetCellText(sheetName, headerRow + 1, column) ?? string.Empty
                    : string.Empty;

                if (!string.IsNullOrWhiteSpace(topText))
                {
                    currentParent = topText;
                }

                var match = FindMatch(lookup, topText, bottomText, currentParent, binding.HeaderRowCount);
                if (match == null)
                {
                    continue;
                }

                match.ColumnIndex = column;
                result.Add(match);
            }

            return result.ToArray();
        }

        private HeaderLookup BuildLookup(
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings)
        {
            var singleHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
            var activityHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);

            foreach (var mapping in mappings)
            {
                if (mapping == null)
                {
                    continue;
                }

                var headerType = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.HeaderType);
                var apiFieldKey = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.ApiFieldKey);
                var currentSingle = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentSingleHeaderText);
                var currentParentText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentParentHeaderText);
                var currentChildText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentChildHeaderText);
                var template = new WorksheetRuntimeColumn
                {
                    ApiFieldKey = apiFieldKey,
                    HeaderType = headerType,
                    DisplayText = currentSingle,
                    ParentDisplayText = currentParentText,
                    ChildDisplayText = currentChildText,
                    IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
                };

                if (IsSingleHeader(headerType))
                {
                    if (!string.IsNullOrWhiteSpace(currentSingle) && !singleHeaders.ContainsKey(currentSingle))
                    {
                        singleHeaders[currentSingle] = template;
                    }

                    continue;
                }

                if (string.Equals(headerType, "activityProperty", StringComparison.OrdinalIgnoreCase))
                {
                    var activityKey = BuildActivityKey(currentParentText, currentChildText);
                    if (!string.IsNullOrWhiteSpace(activityKey) && !activityHeaders.ContainsKey(activityKey))
                    {
                        activityHeaders[activityKey] = template;
                    }
                }
            }

            return new HeaderLookup(singleHeaders, activityHeaders);
        }

        private WorksheetRuntimeColumn FindMatch(
            HeaderLookup lookup,
            string topText,
            string bottomText,
            string currentParent,
            int headerRowCount)
        {
            if (headerRowCount <= 1)
            {
                if (lookup.SingleHeaders.TryGetValue(topText, out var singleHeader))
                {
                    return CloneSingleHeader(singleHeader);
                }

                return null;
            }

            if (lookup.SingleHeaders.TryGetValue(topText, out var mergedSingleHeader) &&
                (string.IsNullOrWhiteSpace(bottomText) || string.Equals(bottomText, topText, StringComparison.Ordinal)))
            {
                return CloneSingleHeader(mergedSingleHeader);
            }

            var activityLookupKey = BuildActivityKey(currentParent, bottomText);
            if (lookup.ActivityHeaders.TryGetValue(activityLookupKey, out var activityHeader))
            {
                return CloneActivityHeader(activityHeader);
            }

            return null;
        }

        private static WorksheetRuntimeColumn CloneSingleHeader(WorksheetRuntimeColumn template)
        {
            return new WorksheetRuntimeColumn
            {
                ApiFieldKey = template.ApiFieldKey,
                HeaderType = template.HeaderType,
                DisplayText = template.DisplayText,
                ParentDisplayText = string.Empty,
                ChildDisplayText = string.Empty,
                IsIdColumn = template.IsIdColumn,
            };
        }

        private static WorksheetRuntimeColumn CloneActivityHeader(WorksheetRuntimeColumn template)
        {
            return new WorksheetRuntimeColumn
            {
                ApiFieldKey = template.ApiFieldKey,
                HeaderType = template.HeaderType,
                DisplayText = template.ChildDisplayText,
                ParentDisplayText = template.ParentDisplayText,
                ChildDisplayText = template.ChildDisplayText,
                IsIdColumn = template.IsIdColumn,
            };
        }

        private static string BuildActivityKey(string parentText, string childText)
        {
            if (string.IsNullOrWhiteSpace(parentText) || string.IsNullOrWhiteSpace(childText))
            {
                return string.Empty;
            }

            return parentText + "\u001f" + childText;
        }

        private sealed class HeaderLookup
        {
            public HeaderLookup(
                IReadOnlyDictionary<string, WorksheetRuntimeColumn> singleHeaders,
                IReadOnlyDictionary<string, WorksheetRuntimeColumn> activityHeaders)
            {
                SingleHeaders = singleHeaders ?? throw new ArgumentNullException(nameof(singleHeaders));
                ActivityHeaders = activityHeaders ?? throw new ArgumentNullException(nameof(activityHeaders));
            }

            public IReadOnlyDictionary<string, WorksheetRuntimeColumn> SingleHeaders { get; }

            public IReadOnlyDictionary<string, WorksheetRuntimeColumn> ActivityHeaders { get; }
        }

        private static bool IsSingleHeader(string headerType)
        {
            return string.IsNullOrWhiteSpace(headerType) ||
                   string.Equals(headerType, "single", StringComparison.OrdinalIgnoreCase);
        }
    }
}
