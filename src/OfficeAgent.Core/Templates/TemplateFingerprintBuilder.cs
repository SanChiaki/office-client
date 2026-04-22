using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Templates
{
    public sealed class TemplateFingerprintBuilder
    {
        public string Build(TemplateDefinition template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            var canonicalRows = (template.FieldMappings ?? Array.Empty<TemplateFieldMappingRow>())
                .Select(BuildCanonicalRow)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();

            var payload = string.Join(
                "\n",
                new[]
                {
                    template.SystemKey ?? string.Empty,
                    template.ProjectId ?? string.Empty,
                    template.ProjectName ?? string.Empty,
                    template.HeaderStartRow.ToString(),
                    template.HeaderRowCount.ToString(),
                    template.DataStartRow.ToString(),
                    BuildFieldMappingDefinitionFingerprint(template.FieldMappingDefinition),
                }.Concat(canonicalRows));

            return ComputeSha256Hex(payload);
        }

        public static string BuildFieldMappingDefinitionFingerprint(FieldMappingTableDefinition definition)
        {
            if (definition == null)
            {
                return ComputeSha256Hex(string.Empty);
            }

            var columns = definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>();
            var canonicalColumns = columns
                .Select((column, index) =>
                {
                    var value = column ?? new FieldMappingColumnDefinition();
                    return string.Join(
                        "|",
                        index.ToString(),
                        value.ColumnName ?? string.Empty,
                        value.Role.ToString(),
                        value.RoleKey ?? string.Empty);
                })
                .ToArray();

            var payload = string.Join(
                "\n",
                new[] { definition.SystemKey ?? string.Empty }.Concat(canonicalColumns));

            return ComputeSha256Hex(payload);
        }

        private static string BuildCanonicalRow(TemplateFieldMappingRow row)
        {
            var values = row?.Values ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var pairs = values
                .Where(pair => !string.Equals(pair.Key, "SheetName", StringComparison.OrdinalIgnoreCase))
                .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .Select(pair => string.Join("=", pair.Key ?? string.Empty, pair.Value ?? string.Empty));

            return string.Join("|", pairs);
        }

        private static string ComputeSha256Hex(string value)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
                var hash = sha.ComputeHash(bytes);
                var builder = new StringBuilder(hash.Length * 2);
                foreach (var item in hash)
                {
                    builder.Append(item.ToString("x2"));
                }

                return builder.ToString();
            }
        }
    }
}
