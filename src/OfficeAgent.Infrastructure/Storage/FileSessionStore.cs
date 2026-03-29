using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Infrastructure.Storage
{
    public sealed class FileSessionStore
    {
        private readonly string storageDirectory;

        public FileSessionStore(string storageDirectory)
        {
            this.storageDirectory = storageDirectory ?? throw new ArgumentNullException(nameof(storageDirectory));
        }

        public SessionState Load()
        {
            var sessionsPath = GetSessionsPath();
            if (!File.Exists(sessionsPath))
            {
                return CreateDefaultState();
            }

            try
            {
                var state = JsonConvert.DeserializeObject<SessionState>(File.ReadAllText(sessionsPath));
                return Normalize(state);
            }
            catch (JsonException)
            {
                return CreateDefaultState();
            }
        }

        public void Save(SessionState state)
        {
            Directory.CreateDirectory(storageDirectory);
            var normalizedState = Normalize(state);
            File.WriteAllText(GetSessionsPath(), JsonConvert.SerializeObject(normalizedState, Formatting.Indented));
        }

        private string GetSessionsPath()
        {
            return Path.Combine(storageDirectory, "sessions.json");
        }

        private static SessionState Normalize(SessionState state)
        {
            if (state == null || state.Sessions == null || state.Sessions.Count == 0)
            {
                return CreateDefaultState();
            }

            var sanitizedSessions = state.Sessions
                .Where(session => session != null && !string.IsNullOrWhiteSpace(session.Id))
                .ToList();

            if (sanitizedSessions.Count == 0)
            {
                return CreateDefaultState();
            }

            foreach (var session in sanitizedSessions)
            {
                session.Messages = session.Messages?
                    .Where(message => message != null)
                    .ToList() ?? new List<ChatMessage>();
            }

            state.Sessions = sanitizedSessions;

            if (string.IsNullOrWhiteSpace(state.ActiveSessionId) || state.Sessions.Find(session => session.Id == state.ActiveSessionId) == null)
            {
                state.ActiveSessionId = state.Sessions[0].Id;
            }

            return state;
        }

        private static SessionState CreateDefaultState()
        {
            var sessionId = $"session-{Guid.NewGuid():N}";
            var timestamp = DateTime.UtcNow;

            return new SessionState
            {
                ActiveSessionId = sessionId,
                Sessions =
                {
                    new ChatSession
                    {
                        Id = sessionId,
                        Title = "New chat",
                        CreatedAtUtc = timestamp,
                        UpdatedAtUtc = timestamp,
                    },
                },
            };
        }
    }
}
