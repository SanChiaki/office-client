using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Skills;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class UploadDataSkillTests
    {
        private const string UploadToProjectA = "\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u5230\u9879\u76EEA";
        private const string ProjectA = "\u9879\u76EEA";

        [Fact]
        public void ExecuteBuildsAnUploadPreviewFromTheCurrentSelection()
        {
            var executor = new FakeExcelCommandExecutor();
            var gateway = new FakeUploadDataGateway();
            var skill = new UploadDataSkill(executor, gateway);

            var result = skill.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadToProjectA,
                Confirmed = false,
            });

            Assert.Equal(1, executor.ExecuteCalls);
            Assert.Equal(0, gateway.UploadCalls);
            Assert.Equal(SkillNames.UploadData, result.SkillName);
            Assert.True(result.RequiresConfirmation);
            Assert.Equal("preview", result.Status);
            Assert.NotNull(result.UploadPreview);
            Assert.Equal(ProjectA, result.UploadPreview.ProjectName);
            Assert.Equal(new[] { "Name", "Region" }, result.UploadPreview.Headers);
            Assert.Equal("Project A", result.UploadPreview.Records[0]["Name"]);
            Assert.Equal("CN", result.UploadPreview.Records[0]["Region"]);
        }

        [Fact]
        public void ExecuteUploadsTheProvidedPreviewAfterConfirmation()
        {
            var executor = new FakeExcelCommandExecutor();
            var gateway = new FakeUploadDataGateway();
            var skill = new UploadDataSkill(executor, gateway);
            var preview = new UploadPreview
            {
                ProjectName = ProjectA,
                SheetName = "Sheet1",
                Address = "A1:C3",
                Headers = new[] { "Name", "Region" },
                Rows = new[]
                {
                    new[] { "Project A", "CN" },
                    new[] { "Project B", "US" },
                },
                Records = new[]
                {
                    new Dictionary<string, string>
                    {
                        ["Name"] = "Project A",
                        ["Region"] = "CN",
                    },
                    new Dictionary<string, string>
                    {
                        ["Name"] = "Project B",
                        ["Region"] = "US",
                    },
                },
            };

            var result = skill.Execute(new AgentCommandEnvelope
            {
                SkillName = SkillNames.UploadData,
                UserInput = $"/upload_data {UploadToProjectA}",
                Confirmed = true,
                UploadPreview = preview,
            });

            Assert.Equal(0, executor.ExecuteCalls);
            Assert.Equal(1, gateway.UploadCalls);
            Assert.Same(preview, gateway.LastPreview);
            Assert.False(result.RequiresConfirmation);
            Assert.Equal("completed", result.Status);
            Assert.Contains("Uploaded 2 row(s)", result.Message);
        }

        [Fact]
        public void ExecuteRejectsMalformedConfirmedPreviewPayloads()
        {
            var executor = new FakeExcelCommandExecutor();
            var gateway = new FakeUploadDataGateway();
            var skill = new UploadDataSkill(executor, gateway);
            var malformedPreview = new UploadPreview
            {
                ProjectName = ProjectA,
                SheetName = "Sheet1",
                Address = "A1:C3",
                Headers = null,
                Rows = null,
                Records = null,
            };

            var error = Assert.Throws<ArgumentException>(() => skill.Execute(new AgentCommandEnvelope
            {
                SkillName = SkillNames.UploadData,
                UserInput = $"/upload_data {UploadToProjectA}",
                Confirmed = true,
                UploadPreview = malformedPreview,
            }));

            Assert.Equal(0, executor.ExecuteCalls);
            Assert.Equal(0, gateway.UploadCalls);
            Assert.Contains("complete preview payload", error.Message);
        }

        private sealed class FakeExcelCommandExecutor : IExcelCommandExecutor
        {
            public int ExecuteCalls { get; private set; }

            public ExcelCommandResult Preview(ExcelCommand command)
            {
                throw new System.NotSupportedException();
            }

            public ExcelCommandResult Execute(ExcelCommand command)
            {
                ExecuteCalls++;

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
            public int UploadCalls { get; private set; }

            public UploadPreview LastPreview { get; private set; }

            public UploadExecutionResult Upload(UploadPreview preview)
            {
                UploadCalls++;
                LastPreview = preview;

                return new UploadExecutionResult
                {
                    SavedCount = preview?.Records?.Length ?? 0,
                    Message = $"Uploaded {preview?.Records?.Length ?? 0} row(s) to {preview?.ProjectName}.",
                };
            }
        }
    }
}
