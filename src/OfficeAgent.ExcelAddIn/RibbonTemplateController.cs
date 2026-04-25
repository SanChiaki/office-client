using System;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class RibbonTemplateController
    {
        private readonly ITemplateCatalog templateCatalog;
        private readonly Func<string> activeSheetNameProvider;
        private readonly IRibbonTemplateDialogService dialogService;
        private string lastRefreshedSheetName = string.Empty;

        public RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider)
            : this(templateCatalog, activeSheetNameProvider, new RibbonTemplateDialogService())
        {
        }

        internal RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider,
            IRibbonTemplateDialogService dialogService)
        {
            this.templateCatalog = templateCatalog ?? throw new ArgumentNullException(nameof(templateCatalog));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));

            ActiveTemplateDisplayName = GetStrings().DefaultTemplateDisplayName;
        }

        public event EventHandler TemplateStateChanged;

        public string ActiveTemplateDisplayName { get; private set; }

        public bool CanApplyTemplate { get; private set; }

        public bool CanSaveTemplate { get; private set; }

        public bool CanSaveAsTemplate { get; private set; }

        public void RefreshActiveTemplateStateFromSheetMetadata()
        {
            RefreshTemplateState(activeSheetNameProvider.Invoke() ?? string.Empty);
        }

        internal void RefreshTemplateState(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                lastRefreshedSheetName = string.Empty;
                ApplyState(new SheetTemplateState());
                return;
            }

            if (string.Equals(lastRefreshedSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            lastRefreshedSheetName = sheetName;
            ApplyState(templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState());
        }

        internal void InvalidateRefreshState()
        {
            lastRefreshedSheetName = string.Empty;
        }

        public void ExecuteApplyTemplate()
        {
            try
            {
                var sheetName = GetRequiredSheetName();
                var state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                if (!state.CanApplyTemplate)
                {
                    dialogService.ShowWarning(GetStrings().ProjectSelectionRequiredMessage);
                    return;
                }

                var templates = templateCatalog.ListTemplates(sheetName) ?? Array.Empty<TemplateDefinition>();
                if (templates.Count == 0)
                {
                    dialogService.ShowWarning(GetStrings().TemplateNoAvailableMessage);
                    return;
                }

                var templateId = dialogService.ShowTemplatePicker(state.ProjectDisplayName, templates);
                if (string.IsNullOrWhiteSpace(templateId))
                {
                    return;
                }

                var selectedTemplate = templates.FirstOrDefault(template =>
                    string.Equals(template.TemplateId, templateId, StringComparison.Ordinal));
                if (selectedTemplate == null)
                {
                    dialogService.ShowWarning(GetStrings().TemplateNotFoundMessage);
                    return;
                }

                if (state.IsDirty && !dialogService.ConfirmApplyTemplateOverwrite(selectedTemplate.TemplateName))
                {
                    return;
                }

                templateCatalog.ApplyTemplateToSheet(sheetName, templateId);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo(GetStrings().ApplyTemplateCompletedMessage(selectedTemplate.TemplateName));
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveTemplate()
        {
            try
            {
                var sheetName = GetRequiredSheetName();
                var state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                if (!state.CanSaveTemplate || string.IsNullOrWhiteSpace(state.TemplateId) || !state.TemplateRevision.HasValue)
                {
                    dialogService.ShowWarning(GetStrings().TemplateNoSavableMessage);
                    return;
                }

                if (TrySaveTemplate(sheetName, state, overwriteRevisionConflict: false, GetStrings().SaveTemplateCompletedMessage(state.TemplateName)))
                {
                    return;
                }

                var conflictResult = dialogService.ShowTemplateRevisionConflictDialog(
                    state.TemplateName,
                    state.TemplateRevision.Value,
                    state.StoredTemplateRevision ?? state.TemplateRevision.Value);

                if (conflictResult == DialogResult.Yes)
                {
                    TrySaveTemplate(sheetName, state, overwriteRevisionConflict: true, GetStrings().OverwriteTemplateCompletedMessage(state.TemplateName));
                    return;
                }

                if (conflictResult == DialogResult.No)
                {
                    ExecuteSaveAsTemplate();
                }
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveAsTemplate()
        {
            try
            {
                var sheetName = GetRequiredSheetName();
                var state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                if (!state.CanSaveAsTemplate)
                {
                    dialogService.ShowWarning(GetStrings().ProjectSelectionRequiredMessage);
                    return;
                }

                var suggestedTemplateName = GetStrings().FormatSuggestedTemplateCopyName(state.TemplateName);
                var templateName = dialogService.ShowSaveAsTemplateDialog(suggestedTemplateName);
                if (string.IsNullOrWhiteSpace(templateName))
                {
                    return;
                }

                templateCatalog.SaveSheetAsNewTemplate(sheetName, templateName);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo(GetStrings().SaveAsTemplateCompletedMessage(templateName));
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        private bool TrySaveTemplate(
            string sheetName,
            SheetTemplateState state,
            bool overwriteRevisionConflict,
            string successMessage)
        {
            try
            {
                templateCatalog.SaveSheetToExistingTemplate(
                    sheetName,
                    state.TemplateId,
                    state.TemplateRevision ?? 0,
                    overwriteRevisionConflict);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo(successMessage);
                return true;
            }
            catch (InvalidOperationException ex) when (!overwriteRevisionConflict && IsRevisionConflict(ex.Message))
            {
                return false;
            }
        }

        private void ApplyState(SheetTemplateState state)
        {
            CanApplyTemplate = state?.CanApplyTemplate == true;
            CanSaveTemplate = state?.CanSaveTemplate == true;
            CanSaveAsTemplate = state?.CanSaveAsTemplate == true;
            ActiveTemplateDisplayName = string.IsNullOrWhiteSpace(state?.TemplateName)
                ? GetStrings().DefaultTemplateDisplayName
                : state.TemplateName;
            TemplateStateChanged?.Invoke(this, EventArgs.Empty);
        }

        private string GetRequiredSheetName()
        {
            var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new InvalidOperationException("Active worksheet is not available.");
            }

            return sheetName;
        }

        private static bool IsRevisionConflict(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            return string.Equals(message, "模板版本已变化。", StringComparison.Ordinal) ||
                   message.IndexOf("revision", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static HostLocalizedStrings GetStrings()
        {
            return Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
        }
    }
}
