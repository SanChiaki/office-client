using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplateNameDialog : Form
    {
        private readonly TextBox templateNameTextBox;

        public TemplateNameDialog(string suggestedTemplateName)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");

            Text = strings.TemplateNameDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(460, 190);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle());

            var promptLabel = new Label
            {
                AutoSize = true,
                Margin = Padding.Empty,
                MaximumSize = new Size(412, 0),
                Text = strings.TemplateNameDialogPrompt,
            };

            templateNameTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 12),
                Text = suggestedTemplateName ?? string.Empty,
            };

            var okButton = new Button { Text = strings.OkButtonText, Width = 88, Height = 30 };
            okButton.Click += OkButton_Click;
            var cancelButton = new Button { Text = strings.CancelButtonText, Width = 88, Height = 30, DialogResult = DialogResult.Cancel };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);

            root.Controls.Add(promptLabel, 0, 0);
            root.Controls.Add(templateNameTextBox, 0, 1);
            root.Controls.Add(buttons, 0, 2);

            Controls.Add(root);
            AcceptButton = okButton;
            CancelButton = cancelButton;
            Shown += (sender, args) =>
            {
                templateNameTextBox.Focus();
                templateNameTextBox.SelectAll();
            };
        }

        public string TemplateName { get; private set; } = string.Empty;

        private void OkButton_Click(object sender, EventArgs e)
        {
            var value = (templateNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
                OperationResultDialog.ShowWarning(strings.TemplateNameRequiredMessage);
                return;
            }

            TemplateName = value;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
