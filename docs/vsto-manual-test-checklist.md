# OfficeAgent VSTO Manual Test Checklist

## Installer

- Confirm `installer/OfficeAgent.SetupBundle/prereqs/` contains `vstor_redist.exe`, `MicrosoftEdgeWebView2RuntimeInstallerX86.exe`, and `MicrosoftEdgeWebView2RuntimeInstallerX64.exe` before starting the build.
- Run `installer/OfficeAgent.Setup/build.ps1` and confirm `artifacts/installer/OfficeAgent.Setup.exe`, `artifacts/installer/OfficeAgent.Setup-x86.msi`, and `artifacts/installer/OfficeAgent.Setup-x64.msi` are created.
- Run `OfficeAgent.Setup.exe` on a machine missing both prerequisites and confirm it installs VSTO Runtime, installs WebView2 Runtime, and then installs OfficeAgent.
- Run `OfficeAgent.Setup.exe` on a machine with both prerequisites already installed and confirm it skips both prerequisite installers.
- Run `OfficeAgent.Setup.exe` twice on the same machine and confirm the second run does not reinstall prerequisites and falls through to normal OfficeAgent maintenance behavior.
- Choose the MSI that matches the target Excel bitness only for direct enterprise distribution or debugging. Do not install the x86 package for x64 Excel or the x64 package for x86 Excel.
- Install the direct MSI under a standard user profile.
- Confirm files are deployed under `%LocalAppData%\\OfficeAgent\\ExcelAddIn`.
- Confirm Excel add-in registry entries exist under `HKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\OfficeAgent.ExcelAddIn`.
- On a machine missing the VSTO runtime, confirm the direct MSI blocks with a clear prerequisite message.
- On a machine missing the WebView2 runtime, confirm the direct MSI blocks with a clear prerequisite message.
- Note the current MVP manifests are signed with the development publisher `OfficeAgent Dev Certificate`; for distribution outside the build machine, replace it with a trusted code-signing certificate or import the publisher certificate through your enterprise deployment flow.

## Excel Startup

- Start Excel 2019 x86 on Windows with the x86 MSI and confirm `OfficeAgent` loads without manual sideload.
- Start Excel 2019 x64 on Windows with the x64 MSI and confirm `OfficeAgent` loads without manual sideload.
- Close and reopen Excel and confirm the add-in still loads automatically.

## Task Pane

- Use the Ribbon button to open and close the task pane repeatedly.
- Confirm the task pane does not duplicate after repeated toggles.
- Confirm the WebView2 missing-runtime fallback message appears if WebView2 is not installed.

## Session And Settings

- Open Settings and save `API Key`, `Base URL`, `Business Base URL`, `Model`, `SSO URL`, and `登录成功路径`.
- Confirm `Base URL` stays reserved for the LLM endpoint and `Business Base URL` points to the business API or mock server.
- Restart Excel and confirm settings reload correctly.
- Create or switch sessions and confirm existing thread history is preserved per session.

## Selection Context

- Select a contiguous range and confirm workbook, sheet, address, row count, column count, and headers update.
- Select a non-contiguous range and confirm the warning is shown.

## upload_data

- Trigger `upload_data` with natural language and confirm a preview card appears.
- Trigger `upload_data` with `/upload_data ...` and confirm it also routes to the skill.
- Cancel the preview and confirm the thread logs the cancellation without changing Excel.
- Confirm the preview and verify the external API is called with the selected rows.
- Simulate a 4xx/5xx API failure and confirm the error message is shown in the task pane.
- Configure a `Business Base URL` with a path prefix such as `/v1/` and confirm the request preserves the prefix.

## Excel Command Confirmation

- Run a read command and confirm it executes immediately.
- Run a write command and confirm it requires preview + confirmation.
- Leave a confirmation card open and verify the composer stays disabled until confirm or cancel.

## Ribbon Sync

- Bind a blank worksheet through the Ribbon project dropdown and confirm the layout dialog appears with defaults `HeaderStartRow = 1`, `HeaderRowCount = 2`, `DataStartRow = 3`.
- Confirm the layout dialog and enter custom values, then verify `AI_Setting` writes one `SheetBindings` row with the user-entered layout values.
- Confirming project selection should still not auto-initialize the current sheet; `SheetFieldMappings` remains unchanged until `初始化当前表` is clicked.
- Open `AI_Setting` and confirm it uses one worksheet with two readable sections: `SheetBindings` on top, `SheetFieldMappings` below, each with a title row, a header row, and data rows.
- Confirm there are two blank separator rows between `SheetBindings` and `SheetFieldMappings`, and that metadata is no longer stored as flattened `tableName + values` rows.
- Switch to a worksheet with existing binding metadata and confirm the Ribbon dropdown automatically rehydrates that project as `ProjectId-DisplayName` instead of showing `先选择项目`.
- Save a workbook with `AI_Setting` as the active sheet, reopen Excel from the desktop shortcut, and confirm the Ribbon dropdown shows `先选择项目` unless `SheetBindings` 里存在 `AI_Setting` 这条显式绑定记录。
- Switch from a bound business sheet to `AI_Setting` and confirm the Ribbon dropdown clears back to `先选择项目` when `AI_Setting` itself has no binding.
- Switch to a worksheet without binding metadata and confirm the Ribbon dropdown shows `先选择项目`.
- On a sheet that already has binding metadata, switch to another project and confirm the layout dialog defaults reuse the current sheet's saved layout values.
- Reselect the already bound project (`same systemKey + projectId`) and confirm no layout dialog appears and `SheetBindings` is not rewritten.
- Cancel the layout dialog while switching projects and confirm both `AI_Setting` binding data and Ribbon dropdown project stay unchanged.
- After switching to another project and confirming the dialog, verify old `SheetFieldMappings` are cleared; before clicking `初始化当前表`, running download/upload should report that the current sheet is not initialized.
- Enter invalid values in the layout dialog (for example overlaps between header/data regions) and confirm validation error is shown while keeping the dialog open.
- Start Excel while unauthenticated against a protected project API and confirm the project dropdown shows `请先登录`.
- Configure the project API to return an empty array and confirm the project dropdown shows `无可用项目`.
- Click `初始化当前表` on a sheet that already contains business cells and confirm only `AI_Setting` changes; the business area should remain untouched.
- Click `全量下载`, `部分下载`, and `部分上传` and confirm each action uses a native Office/WinForms confirmation dialog instead of the task pane.
- Confirm the `下载` and `上传` controls are rendered in separate Ribbon groups, that the upload group only shows `部分上传`, and that there is no `全量上传` or `增量上传` button.
- Edit `AI_Setting` so `HeaderStartRow = 3`, `HeaderRowCount = 2`, and `DataStartRow = 6`, then run `全量下载` and confirm headers/data are written at the configured rows.
- On a sheet that already has recognizable headers, run `全量下载` and confirm the plugin refreshes data cells without rewriting those existing headers.
- Modify a current display name in `SheetFieldMappings`, update the matching Excel header text manually, then run `部分下载` or `部分上传` and confirm the column still resolves by current header text.
- Verify the task pane button and login button still work after the Ribbon Sync controls are added.
