# OfficeAgent VSTO Manual Test Checklist

## Installer

- Run `installer/OfficeAgent.Setup/build.ps1` and confirm both `artifacts/installer/OfficeAgent.Setup-x86.msi` and `artifacts/installer/OfficeAgent.Setup-x64.msi` are created.
- Choose the MSI that matches the target Excel bitness. Do not install the x86 package for x64 Excel or the x64 package for x86 Excel.
- Install the MSI under a standard user profile.
- Confirm files are deployed under `%LocalAppData%\\OfficeAgent\\ExcelAddIn`.
- Confirm Excel add-in registry entries exist under `HKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\OfficeAgent.ExcelAddIn`.
- On a machine missing the VSTO runtime, confirm the installer blocks with a clear prerequisite message.
- On a machine missing the WebView2 runtime, confirm the installer blocks with a clear prerequisite message.
- Confirm the current MVP deployment flow expects WebView2 runtime preinstallation; the MSI does not bootstrap the runtime yet.
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

- Open Settings and save `API Key`, `Base URL`, and `Model`.
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
- Configure a `Base URL` with a path prefix such as `/v1/` and confirm the request preserves the prefix.

## Excel Command Confirmation

- Run a read command and confirm it executes immediately.
- Run a write command and confirm it requires preview + confirmation.
- Leave a confirmation card open and verify the composer stays disabled until confirm or cancel.
