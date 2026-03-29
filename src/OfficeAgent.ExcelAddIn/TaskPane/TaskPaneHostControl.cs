using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using OfficeAgent.ExcelAddIn.WebBridge;

namespace OfficeAgent.ExcelAddIn.TaskPane
{
    internal sealed class TaskPaneHostControl : UserControl
    {
        private readonly WebView2 webView;
        private readonly WebViewBootstrapper bootstrapper;
        private bool isInitialized;

        public TaskPaneHostControl()
        {
            Dock = DockStyle.Fill;

            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(webView);

            bootstrapper = new WebViewBootstrapper(webView);
            Load += TaskPaneHostControl_Load;
        }

        private async void TaskPaneHostControl_Load(object sender, EventArgs e)
        {
            if (isInitialized || DesignMode)
            {
                return;
            }

            isInitialized = true;

            try
            {
                await bootstrapper.InitializeAsync();
            }
            catch (WebView2RuntimeNotFoundException)
            {
                webView.Visible = false;
                Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    Text = "WebView2 Runtime is required to render OfficeAgent.",
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                });
            }
        }
    }
}
