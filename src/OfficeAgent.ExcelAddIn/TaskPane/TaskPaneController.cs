using Microsoft.Office.Core;

namespace OfficeAgent.ExcelAddIn.TaskPane
{
    internal sealed class TaskPaneController
    {
        private readonly ThisAddIn addIn;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;
        private TaskPaneHostControl hostControl;

        public TaskPaneController(ThisAddIn addIn)
        {
            this.addIn = addIn;
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

            hostControl = new TaskPaneHostControl();
            taskPane = addIn.CustomTaskPanes.Add(hostControl, "OfficeAgent");
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 420;
            taskPane.Visible = false;
        }
    }
}
