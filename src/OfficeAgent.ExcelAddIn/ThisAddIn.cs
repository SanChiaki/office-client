using System;
using OfficeAgent.ExcelAddIn.TaskPane;

namespace OfficeAgent.ExcelAddIn
{
    public partial class ThisAddIn
    {
        internal TaskPaneController TaskPaneController { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            TaskPaneController = new TaskPaneController(this);
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
