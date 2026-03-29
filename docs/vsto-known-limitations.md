# OfficeAgent VSTO MVP Known Limitations

## Distribution

- The current MSI does not bootstrap the WebView2 runtime. Deploy WebView2 separately before the add-in MSI.
- The current VSTO manifests are signed with `OfficeAgent Dev Certificate`. Replace it with a trusted enterprise or public code-signing certificate before broad rollout.

## Excel behavior

- Non-contiguous selections are blocked.
- Merged cells are rejected for table reads and write-range commands.
- Protected worksheets and protected workbook structure block write operations such as write-range, add-sheet, rename-sheet, and delete-sheet.
- General chat is still a scaffold. The usable business workflow in this MVP is `upload_data`, plus the direct Excel command bridge.

## Verification scope

- Automated tests and packaging build were run on the current development machine.
- Manual QA in supported Excel desktop environments still needs to be completed from `docs/vsto-manual-test-checklist.md`.
