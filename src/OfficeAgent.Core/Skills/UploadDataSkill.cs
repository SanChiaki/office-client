using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Skills
{
    public sealed class UploadDataSkill : IAgentSkill
    {
        private readonly IExcelCommandExecutor excelCommandExecutor;
        private readonly IUploadDataGateway uploadDataGateway;

        public UploadDataSkill(IExcelCommandExecutor excelCommandExecutor, IUploadDataGateway uploadDataGateway)
        {
            this.excelCommandExecutor = excelCommandExecutor ?? throw new ArgumentNullException(nameof(excelCommandExecutor));
            this.uploadDataGateway = uploadDataGateway ?? throw new ArgumentNullException(nameof(uploadDataGateway));
        }

        public string SkillName => SkillNames.UploadData;

        public AgentCommandResult Execute(AgentCommandEnvelope envelope)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException(nameof(envelope));
            }

            if (envelope.Confirmed)
            {
                if (envelope.UploadPreview == null)
                {
                    throw new ArgumentException("upload_data confirmation requires a preview payload.");
                }

                ValidateConfirmedPreview(envelope.UploadPreview);
                var uploadResult = uploadDataGateway.Upload(envelope.UploadPreview);
                return new AgentCommandResult
                {
                    Route = AgentRouteTypes.Skill,
                    SkillName = SkillName,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = string.IsNullOrWhiteSpace(uploadResult.Message)
                        ? $"Uploaded {uploadResult.SavedCount} row(s) to {envelope.UploadPreview.ProjectName}."
                        : uploadResult.Message,
                    UploadPreview = envelope.UploadPreview,
                };
            }

            var selectionResult = excelCommandExecutor.Execute(new ExcelCommand
            {
                CommandType = ExcelCommandTypes.ReadSelectionTable,
                Confirmed = false,
            });
            var table = selectionResult.Table ?? throw new InvalidOperationException("The current Excel selection does not contain tabular data.");
            if (table.Headers == null || table.Headers.Length == 0)
            {
                throw new InvalidOperationException("upload_data requires a selection with header cells in the first row.");
            }

            var preview = BuildUploadPreview(ExtractProjectName(envelope.UserInput), table);
            return new AgentCommandResult
            {
                Route = AgentRouteTypes.Skill,
                SkillName = SkillName,
                RequiresConfirmation = true,
                Status = "preview",
                Message = $"Review the upload payload before sending it to {preview.ProjectName}.",
                Preview = new ExcelCommandPreview
                {
                    Title = "Upload selected data",
                    Summary = $"Upload {preview.Records.Length} row(s) to {preview.ProjectName}",
                    Details = new[]
                    {
                        $"Source: {preview.SheetName}!{preview.Address}",
                        $"Fields: {string.Join(", ", preview.Headers)}",
                    },
                },
                UploadPreview = preview,
            };
        }

        private static UploadPreview BuildUploadPreview(string projectName, ExcelTableData table)
        {
            var records = table.Rows
                .Select((row) => BuildRecord(table.Headers, row))
                .ToArray();

            return new UploadPreview
            {
                ProjectName = projectName,
                SheetName = table.SheetName,
                Address = table.Address,
                Headers = table.Headers,
                Rows = table.Rows,
                Records = records,
            };
        }

        private static Dictionary<string, string> BuildRecord(string[] headers, string[] row)
        {
            var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < headers.Length; index++)
            {
                var header = headers[index]?.Trim();
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                record[header] = index < (row?.Length ?? 0) ? row[index] ?? string.Empty : string.Empty;
            }

            return record;
        }

        private static string ExtractProjectName(string userInput)
        {
            var normalizedInput = (userInput ?? string.Empty).Trim();
            if (normalizedInput.StartsWith("/upload_data", StringComparison.OrdinalIgnoreCase))
            {
                normalizedInput = normalizedInput.Substring("/upload_data".Length).Trim();
            }

            var chineseIndex = normalizedInput.LastIndexOf("上传到", StringComparison.Ordinal);
            if (chineseIndex >= 0)
            {
                var projectName = normalizedInput.Substring(chineseIndex + "上传到".Length).Trim(' ', '。', '.');
                if (!string.IsNullOrWhiteSpace(projectName))
                {
                    return projectName;
                }
            }

            var englishMatch = Regex.Match(normalizedInput, "(?i)\\bto\\s+(.+)$");
            if (englishMatch.Success)
            {
                var projectName = englishMatch.Groups[1].Value.Trim(' ', '.');
                if (!string.IsNullOrWhiteSpace(projectName))
                {
                    return projectName;
                }
            }

            throw new ArgumentException("upload_data requires a target project name.");
        }

        private static void ValidateConfirmedPreview(UploadPreview preview)
        {
            if (preview == null)
            {
                throw new ArgumentException("upload_data confirmation requires a preview payload.");
            }

            if (string.IsNullOrWhiteSpace(preview.ProjectName) ||
                string.IsNullOrWhiteSpace(preview.SheetName) ||
                string.IsNullOrWhiteSpace(preview.Address) ||
                preview.Headers == null ||
                preview.Headers.Length == 0 ||
                preview.Rows == null ||
                preview.Records == null)
            {
                throw new ArgumentException("upload_data confirmation requires a complete preview payload.");
            }
        }
    }
}
