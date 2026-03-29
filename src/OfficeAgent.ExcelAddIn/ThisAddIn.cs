using System;
using System.IO;
using OfficeAgent.ExcelAddIn.TaskPane;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn
{
    public partial class ThisAddIn
    {
        internal TaskPaneController TaskPaneController { get; private set; }
        internal FileSessionStore SessionStore { get; private set; }
        internal FileSettingsStore SettingsStore { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var appDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAgent");
            SessionStore = new FileSessionStore(Path.Combine(appDataDirectory, "sessions"));
            SettingsStore = new FileSettingsStore(
                Path.Combine(appDataDirectory, "settings.json"),
                new DpapiSecretProtector());
            TaskPaneController = new TaskPaneController(this, SessionStore, SettingsStore);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}
