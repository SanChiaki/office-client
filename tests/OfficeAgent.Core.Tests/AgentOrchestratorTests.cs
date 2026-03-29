using OfficeAgent.Core.Models;
using OfficeAgent.Core.Orchestration;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Skills;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class AgentOrchestratorTests
    {
        private const string UploadToProjectA = "\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u5230\u9879\u76EEA";
        private const string ProjectA = "\u9879\u76EEA";
        private const string SummarizeRequest = "\u5E2E\u6211\u603B\u7ED3\u4E00\u4E0B\u8FD9\u4E2A\u5DE5\u4F5C\u7C3F";
        private const string UploadWithoutTargetProject = "upload project data";
        private const string UploadWithoutTargetSeparator = "\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u9879\u76EEA";

        [Fact]
        public void ExecuteRoutesSlashUploadCommandToTheUploadDataSkill()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = $"/upload_data {UploadToProjectA}",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Skill, result.Route);
            Assert.Equal(SkillNames.UploadData, result.SkillName);
            Assert.True(result.RequiresConfirmation);
            Assert.Equal("preview", result.Status);
            Assert.NotNull(result.UploadPreview);
        }

        [Fact]
        public void ExecuteRoutesNaturalLanguageUploadIntentToTheUploadDataSkill()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadToProjectA,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Skill, result.Route);
            Assert.Equal(SkillNames.UploadData, result.SkillName);
            Assert.True(result.RequiresConfirmation);
            Assert.Equal(ProjectA, result.UploadPreview.ProjectName);
        }

        [Fact]
        public void ExecuteReturnsChatFallbackForUnknownUserInput()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = SummarizeRequest,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Contains("not implemented", result.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExecuteReturnsChatFallbackForEnglishUploadTextWithoutATargetProject()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadWithoutTargetProject,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Contains("not implemented", result.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExecuteReturnsChatFallbackForChineseUploadTextWithoutExplicitTargetSeparator()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadWithoutTargetSeparator,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Contains("not implemented", result.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        private static AgentOrchestrator CreateOrchestrator()
        {
            var skill = new UploadDataSkill(
                new FakeExcelCommandExecutor(),
                new FakeUploadDataGateway());

            return new AgentOrchestrator(new SkillRegistry(skill));
        }

        private sealed class FakeExcelCommandExecutor : IExcelCommandExecutor
        {
            public ExcelCommandResult Preview(ExcelCommand command)
            {
                throw new System.NotSupportedException();
            }

            public ExcelCommandResult Execute(ExcelCommand command)
            {
                return new ExcelCommandResult
                {
                    CommandType = ExcelCommandTypes.ReadSelectionTable,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = "Read selection from Sheet1 A1:C3.",
                    Table = new ExcelTableData
                    {
                        SheetName = "Sheet1",
                        Address = "A1:C3",
                        Headers = new[] { "Name", "Region" },
                        Rows = new[]
                        {
                            new[] { "Project A", "CN" },
                            new[] { "Project B", "US" },
                        },
                    },
                };
            }
        }

        private sealed class FakeUploadDataGateway : IUploadDataGateway
        {
            public UploadExecutionResult Upload(UploadPreview preview)
            {
                return new UploadExecutionResult
                {
                    SavedCount = preview?.Records?.Length ?? 0,
                    Message = $"Uploaded {preview?.Records?.Length ?? 0} row(s).",
                };
            }
        }
    }
}
