using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class TemplateOverwriteConfirmDialog
    {
        public static bool Confirm(string templateName)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            var result = TemplatePromptDialog.ShowPrompt(
                strings.TemplateOverwriteConfirmTitle,
                strings.TemplateOverwriteConfirmMessage(templateName),
                MessageBoxIcon.Warning,
                new TemplatePromptDialog.DialogButtonSpec(strings.CancelButtonText, DialogResult.No, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec(strings.TemplateOverwriteButtonText, DialogResult.Yes, isAccept: true));

            return result == DialogResult.Yes;
        }
    }
}
