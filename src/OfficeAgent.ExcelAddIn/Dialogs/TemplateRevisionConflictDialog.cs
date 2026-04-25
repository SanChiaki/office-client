using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class TemplateRevisionConflictDialog
    {
        public static DialogResult ShowDecision(string templateName, int sheetRevision, int storedRevision)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            return TemplatePromptDialog.ShowPrompt(
                strings.TemplateRevisionConflictTitle,
                strings.TemplateRevisionConflictMessage(templateName, sheetRevision, storedRevision),
                MessageBoxIcon.Warning,
                new TemplatePromptDialog.DialogButtonSpec(strings.CancelButtonText, DialogResult.Cancel, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec(strings.TemplateSaveAsNewButtonText, DialogResult.No),
                new TemplatePromptDialog.DialogButtonSpec(strings.TemplateOverwriteOriginalButtonText, DialogResult.Yes, isAccept: true));
        }
    }
}
