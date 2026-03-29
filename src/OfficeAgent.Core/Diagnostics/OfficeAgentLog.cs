using System;

namespace OfficeAgent.Core.Diagnostics
{
    public static class OfficeAgentLog
    {
        private static readonly object SyncRoot = new object();
        private static Action<OfficeAgentLogEntry> sink = _ => { };

        public static void Configure(Action<OfficeAgentLogEntry> sinkAction)
        {
            lock (SyncRoot)
            {
                sink = sinkAction ?? (_ => { });
            }
        }

        public static void Reset()
        {
            Configure(null);
        }

        public static void Info(string component, string eventName, string message, string details = null)
        {
            Write("info", component, eventName, message, details, exception: null);
        }

        public static void Warn(string component, string eventName, string message, string details = null)
        {
            Write("warn", component, eventName, message, details, exception: null);
        }

        public static void Error(string component, string eventName, string message, Exception exception, string details = null)
        {
            Write("error", component, eventName, message, details, exception);
        }

        private static void Write(string level, string component, string eventName, string message, string details, Exception exception)
        {
            Action<OfficeAgentLogEntry> sinkSnapshot;
            lock (SyncRoot)
            {
                sinkSnapshot = sink;
            }

            sinkSnapshot(new OfficeAgentLogEntry
            {
                TimestampUtc = DateTime.UtcNow,
                Level = level ?? string.Empty,
                Component = component ?? string.Empty,
                EventName = eventName ?? string.Empty,
                Message = message ?? string.Empty,
                Details = details ?? string.Empty,
                Exception = exception?.ToString() ?? string.Empty,
            });
        }
    }
}
