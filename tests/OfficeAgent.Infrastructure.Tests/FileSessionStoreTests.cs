using System;
using System.IO;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class FileSessionStoreTests : IDisposable
    {
        private readonly string tempDirectory;

        public FileSessionStoreTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void LoadReturnsDefaultSessionStateWhenStorageDoesNotExist()
        {
            var store = new FileSessionStore(tempDirectory);

            var state = store.Load();

            Assert.NotNull(state.ActiveSessionId);
            Assert.Single(state.Sessions);
            Assert.Equal(state.ActiveSessionId, state.Sessions[0].Id);
        }

        [Fact]
        public void LoadRecoversFromMalformedJson()
        {
            Directory.CreateDirectory(tempDirectory);
            File.WriteAllText(Path.Combine(tempDirectory, "sessions.json"), "{not-json");

            var store = new FileSessionStore(tempDirectory);

            var state = store.Load();

            Assert.Single(state.Sessions);
            Assert.Equal(state.ActiveSessionId, state.Sessions[0].Id);
        }

        [Fact]
        public void LoadRecoversWhenSessionEntriesContainNullValues()
        {
            Directory.CreateDirectory(tempDirectory);
            File.WriteAllText(
                Path.Combine(tempDirectory, "sessions.json"),
                "{\n  \"activeSessionId\": \"session-1\",\n  \"sessions\": [ null ]\n}");

            var store = new FileSessionStore(tempDirectory);

            var state = store.Load();

            Assert.Single(state.Sessions);
            Assert.NotNull(state.Sessions[0]);
            Assert.Equal(state.ActiveSessionId, state.Sessions[0].Id);
        }

        [Fact]
        public void SaveRoundTripsSessionsAndActiveSession()
        {
            var store = new FileSessionStore(tempDirectory);
            var expected = new SessionState
            {
                ActiveSessionId = "session-1",
                Sessions =
                {
                    new ChatSession
                    {
                        Id = "session-1",
                        Title = "Upload data",
                        CreatedAtUtc = new DateTime(2026, 3, 29, 6, 0, 0, DateTimeKind.Utc),
                        UpdatedAtUtc = new DateTime(2026, 3, 29, 6, 5, 0, DateTimeKind.Utc),
                        Messages =
                        {
                            new ChatMessage
                            {
                                Id = "message-1",
                                Role = "user",
                                Content = "Upload the selected rows",
                                CreatedAtUtc = new DateTime(2026, 3, 29, 6, 1, 0, DateTimeKind.Utc),
                            },
                        },
                    },
                },
            };

            store.Save(expected);
            var actual = store.Load();

            Assert.Equal("session-1", actual.ActiveSessionId);
            Assert.Single(actual.Sessions);
            Assert.Equal("Upload data", actual.Sessions[0].Title);
            Assert.Single(actual.Sessions[0].Messages);
            Assert.Equal("Upload the selected rows", actual.Sessions[0].Messages[0].Content);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }
}
