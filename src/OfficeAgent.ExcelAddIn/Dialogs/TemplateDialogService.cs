using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    public interface IRibbonTemplateDialogService
    {
        string ShowTemplatePicker(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates);

        string ShowSaveAsTemplateDialog(string suggestedTemplateName);

        bool ConfirmApplyTemplateOverwrite(string templateName);

        DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision);

        void ShowInfo(string message);

        void ShowWarning(string message);

        void ShowError(string message);
    }

    internal sealed class RibbonTemplateDialogService : IRibbonTemplateDialogService
    {
        public string ShowTemplatePicker(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates)
        {
            return string.Empty;
        }

        public string ShowSaveAsTemplateDialog(string suggestedTemplateName)
        {
            return string.Empty;
        }

        public bool ConfirmApplyTemplateOverwrite(string templateName)
        {
            return MessageBox.Show(
                    $"当前表存在未保存的模板改动，确认用模板“{templateName}”覆盖吗？",
                    "ISDP",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) == DialogResult.Yes;
        }

        public DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision)
        {
            return MessageBox.Show(
                $"模板“{templateName}”已从版本 {sheetRevision} 更新到版本 {storedRevision}。\r\n是：覆盖原模板\r\n否：另存为新模板\r\n取消：终止操作",
                "ISDP",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);
        }

        public void ShowInfo(string message)
        {
            OperationResultDialog.ShowInfo(message);
        }

        public void ShowWarning(string message)
        {
            OperationResultDialog.ShowWarning(message);
        }

        public void ShowError(string message)
        {
            OperationResultDialog.ShowError(message);
        }
    }
}
