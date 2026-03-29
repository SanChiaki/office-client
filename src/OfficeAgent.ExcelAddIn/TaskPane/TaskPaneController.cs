using Microsoft.Office.Core;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.TaskPane
{
    internal sealed class TaskPaneController
    {
        private readonly ThisAddIn addIn;
        private readonly FileSessionStore sessionStore;
        private readonly FileSettingsStore settingsStore;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private TaskPaneHostControl hostControl;

        public TaskPaneController(ThisAddIn addIn, FileSessionStore sessionStore, FileSettingsStore settingsStore)
        {
            this.addIn = addIn;
            this.sessionStore = sessionStore;
            this.settingsStore = settingsStore;
        }

        public void Toggle()
        {
            EnsureCreated();
            taskPane.Visible = !taskPane.Visible;
        }

        public void Show()
        {
            EnsureCreated();
            taskPane.Visible = true;
        }

        private void EnsureCreated()
        {
            if (taskPane != null)
            {
                return;
            }

            hostControl = new TaskPaneHostControl(sessionStore, settingsStore);
            taskPane = addIn.CustomTaskPanes.Add(hostControl, "OfficeAgent");
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 420;
            taskPane.Visible = false;
        }
    }
}
