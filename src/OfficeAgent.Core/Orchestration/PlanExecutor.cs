using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Orchestration
{
    public sealed class PlanExecutor : IPlanExecutor
    {
        private readonly ConfirmationService confirmationService = new ConfirmationService();
        private readonly IExcelCommandExecutor excelCommandExecutor;
        private readonly ISkillRegistry skillRegistry;

        public PlanExecutor(IExcelCommandExecutor excelCommandExecutor, ISkillRegistry skillRegistry)
        {
            this.excelCommandExecutor = excelCommandExecutor ?? throw new ArgumentNullException(nameof(excelCommandExecutor));
            this.skillRegistry = skillRegistry ?? throw new ArgumentNullException(nameof(skillRegistry));
        }

        public PlanExecutionJournal Execute(AgentPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var journalSteps = new List<PlanExecutionJournalStep>();
            var hasFailures = false;
            var errorMessage = string.Empty;

            for (var index = 0; index < plan.Steps.Length; index++)
            {
                var step = plan.Steps[index];

                try
                {
                    if (IsExcelStep(step.Type))
                    {
                        var command = BuildExcelCommand(step);
                        confirmationService.Validate(command);
                        var result = excelCommandExecutor.Execute(command);
                        journalSteps.Add(new PlanExecutionJournalStep
                        {
                            Type = step.Type,
                            Title = BuildStepTitle(step),
                            Status = NormalizeStatus(result.Status),
                            Message = result.Message ?? string.Empty,
                        });
                        continue;
                    }

                    if (string.Equals(step.Type, PlannerStepTypes.UploadData, StringComparison.Ordinal))
                    {
                        ExecuteUploadDataStep(step, journalSteps);
                        continue;
                    }

                    throw new ArgumentException($"Plan step type '{step.Type}' is not supported.");
                }
                catch (Exception error)
                {
                    var stepTitle = BuildStepTitle(step);
                    hasFailures = true;
                    errorMessage = $"{stepTitle}: {error.Message}".Trim();
                    journalSteps.Add(new PlanExecutionJournalStep
                    {
                        Type = step?.Type ?? string.Empty,
                        Title = stepTitle,
                        Status = "failed",
                        ErrorMessage = error.Message,
                    });

                    for (var remaining = index + 1; remaining < plan.Steps.Length; remaining++)
                    {
                        var skippedStep = plan.Steps[remaining];
                        journalSteps.Add(new PlanExecutionJournalStep
                        {
                            Type = skippedStep?.Type ?? string.Empty,
                            Title = BuildStepTitle(skippedStep),
                            Status = "skipped",
                        });
                    }

                    break;
                }
            }

            return new PlanExecutionJournal
            {
                HasFailures = hasFailures,
                ErrorMessage = errorMessage,
                Steps = journalSteps.ToArray(),
            };
        }

        private void ExecuteUploadDataStep(AgentPlanStep step, List<PlanExecutionJournalStep> journalSteps)
        {
            var skill = skillRegistry.Resolve(SkillNames.UploadData);
            var userInput = ResolveUploadUserInput(step.Args);

            var previewResult = skill.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Skill,
                SkillName = SkillNames.UploadData,
                UserInput = userInput,
                Confirmed = false,
            });

            if (previewResult?.UploadPreview == null)
            {
                throw new InvalidOperationException("upload_data execution requires a preview payload.");
            }

            var finalResult = skill.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Skill,
                SkillName = SkillNames.UploadData,
                UserInput = userInput,
                Confirmed = true,
                UploadPreview = previewResult.UploadPreview,
            });

            journalSteps.Add(new PlanExecutionJournalStep
            {
                Type = step.Type,
                Title = BuildStepTitle(step),
                Status = NormalizeStatus(finalResult.Status),
                Message = finalResult.Message ?? string.Empty,
            });
        }

        private static ExcelCommand BuildExcelCommand(AgentPlanStep step)
        {
            var command = step?.Args?.ToObject<ExcelCommand>() ?? new ExcelCommand();
            command.CommandType = step?.Type ?? string.Empty;
            command.Confirmed = true;
            return command;
        }

        private static bool IsExcelStep(string stepType)
        {
            return string.Equals(stepType, ExcelCommandTypes.WriteRange, StringComparison.Ordinal) ||
                   string.Equals(stepType, ExcelCommandTypes.AddWorksheet, StringComparison.Ordinal) ||
                   string.Equals(stepType, ExcelCommandTypes.RenameWorksheet, StringComparison.Ordinal) ||
                   string.Equals(stepType, ExcelCommandTypes.DeleteWorksheet, StringComparison.Ordinal);
        }

        private static string ResolveUploadUserInput(JObject args)
        {
            var userInput = args?["userInput"]?.Value<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(userInput))
            {
                return userInput;
            }

            var projectName = args?["projectName"]?.Value<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(projectName))
            {
                return $"upload selected data to {projectName}";
            }

            throw new ArgumentException("skill.upload_data requires either userInput or projectName.");
        }

        private static string BuildStepTitle(AgentPlanStep step)
        {
            if (step == null)
            {
                return "Unknown step";
            }

            switch (step.Type)
            {
                case ExcelCommandTypes.AddWorksheet:
                    return $"Add worksheet {step.Args?["newSheetName"]?.Value<string>() ?? string.Empty}".Trim();
                case ExcelCommandTypes.RenameWorksheet:
                    return $"Rename worksheet {step.Args?["sheetName"]?.Value<string>() ?? string.Empty} to {step.Args?["newSheetName"]?.Value<string>() ?? string.Empty}".Trim();
                case ExcelCommandTypes.DeleteWorksheet:
                    return $"Delete worksheet {step.Args?["sheetName"]?.Value<string>() ?? string.Empty}".Trim();
                case ExcelCommandTypes.WriteRange:
                    return $"Write range {step.Args?["targetAddress"]?.Value<string>() ?? string.Empty}".Trim();
                case PlannerStepTypes.UploadData:
                    return "Upload selected data";
                default:
                    return step.Type ?? "Unknown step";
            }
        }

        private static string NormalizeStatus(string status)
        {
            return string.IsNullOrWhiteSpace(status) ? "completed" : status;
        }
    }
}
