using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Orchestration;
using OfficeAgent.Core.Services;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class PlanExecutorTests
    {
        [Fact]
        public void ExecuteRunsExcelAndSkillStepsInOrder()
        {
            var excelExecutor = new RecordingExcelCommandExecutor();
            var uploadSkill = new RecordingSkill();
            var executor = new PlanExecutor(excelExecutor, new RecordingSkillRegistry(uploadSkill));

            var journal = executor.Execute(new AgentPlan
            {
                Summary = "Create a summary sheet and upload the selection",
                Steps = new[]
                {
                    new AgentPlanStep
                    {
                        Type = ExcelCommandTypes.AddWorksheet,
                        Args = JObject.FromObject(new
                        {
                            newSheetName = "Summary",
                        }),
                    },
                    new AgentPlanStep
                    {
                        Type = ExcelCommandTypes.WriteRange,
                        Args = JObject.FromObject(new
                        {
                            targetAddress = "Summary!A1:B2",
                            values = new[]
                            {
                                new[] { "Name", "Region" },
                            },
                        }),
                    },
                    new AgentPlanStep
                    {
                        Type = PlannerStepTypes.UploadData,
                        Args = JObject.FromObject(new
                        {
                            userInput = "upload selected data to Project A",
                        }),
                    },
                },
            });

            Assert.False(journal.HasFailures);
            Assert.Equal(3, journal.Steps.Length);
            Assert.Equal("completed", journal.Steps[0].Status);
            Assert.Equal("completed", journal.Steps[1].Status);
            Assert.Equal("completed", journal.Steps[2].Status);
            Assert.Equal(2, excelExecutor.ExecuteCalls.Count);
            Assert.True(excelExecutor.ExecuteCalls[0].Confirmed);
            Assert.Equal(2, uploadSkill.Envelopes.Count);
            Assert.False(uploadSkill.Envelopes[0].Confirmed);
            Assert.True(uploadSkill.Envelopes[1].Confirmed);
            Assert.NotNull(uploadSkill.Envelopes[1].UploadPreview);
        }

        [Fact]
        public void ExecuteStopsAfterTheFirstFailedStepAndMarksTheRestSkipped()
        {
            var excelExecutor = new RecordingExcelCommandExecutor
            {
                FailOnCommandType = ExcelCommandTypes.WriteRange,
            };
            var executor = new PlanExecutor(excelExecutor, new RecordingSkillRegistry(new RecordingSkill()));

            var journal = executor.Execute(new AgentPlan
            {
                Summary = "Create a summary sheet and upload the selection",
                Steps = new[]
                {
                    new AgentPlanStep
                    {
                        Type = ExcelCommandTypes.AddWorksheet,
                        Args = JObject.FromObject(new
                        {
                            newSheetName = "Summary",
                        }),
                    },
                    new AgentPlanStep
                    {
                        Type = ExcelCommandTypes.WriteRange,
                        Args = JObject.FromObject(new
                        {
                            targetAddress = "Summary!A1:B2",
                            values = new[]
                            {
                                new[] { "Name", "Region" },
                            },
                        }),
                    },
                    new AgentPlanStep
                    {
                        Type = PlannerStepTypes.UploadData,
                        Args = JObject.FromObject(new
                        {
                            userInput = "upload selected data to Project A",
                        }),
                    },
                },
            });

            Assert.True(journal.HasFailures);
            Assert.Equal("completed", journal.Steps[0].Status);
            Assert.Equal("failed", journal.Steps[1].Status);
            Assert.Equal("skipped", journal.Steps[2].Status);
            Assert.Contains("write range", journal.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }

        private sealed class RecordingExcelCommandExecutor : IExcelCommandExecutor
        {
            public string FailOnCommandType { get; set; } = string.Empty;

            public List<ExcelCommand> ExecuteCalls { get; } = new List<ExcelCommand>();

            public ExcelCommandResult Preview(ExcelCommand command)
            {
                throw new NotSupportedException();
            }

            public ExcelCommandResult Execute(ExcelCommand command)
            {
                ExecuteCalls.Add(command);
                if (string.Equals(command.CommandType, FailOnCommandType, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException($"Unable to execute {command.CommandType}.");
                }

                return new ExcelCommandResult
                {
                    CommandType = command.CommandType,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = $"Executed {command.CommandType}.",
                };
            }
        }

        private sealed class RecordingSkillRegistry : ISkillRegistry
        {
            private readonly IAgentSkill skill;

            public RecordingSkillRegistry(IAgentSkill skill)
            {
                this.skill = skill;
            }

            public IAgentSkill Resolve(string skillName)
            {
                return skill;
            }
        }

        private sealed class RecordingSkill : IAgentSkill
        {
            public string SkillName => SkillNames.UploadData;

            public List<AgentCommandEnvelope> Envelopes { get; } = new List<AgentCommandEnvelope>();

            public AgentCommandResult Execute(AgentCommandEnvelope envelope)
            {
                Envelopes.Add(envelope);

                if (!envelope.Confirmed)
                {
                    return new AgentCommandResult
                    {
                        Route = AgentRouteTypes.Skill,
                        SkillName = SkillName,
                        RequiresConfirmation = true,
                        Status = "preview",
                        Message = "Preview ready.",
                        UploadPreview = new UploadPreview
                        {
                            ProjectName = "Project A",
                            SheetName = "Sheet1",
                            Address = "A1:B3",
                            Headers = new[] { "Name", "Region" },
                            Rows = new[]
                            {
                                new[] { "Project A", "CN" },
                            },
                            Records = new[]
                            {
                                new Dictionary<string, string>
                                {
                                    ["Name"] = "Project A",
                                    ["Region"] = "CN",
                                },
                            },
                        },
                    };
                }

                return new AgentCommandResult
                {
                    Route = AgentRouteTypes.Skill,
                    SkillName = SkillName,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = "Uploaded 1 row.",
                };
            }
        }
    }
}
