# OfficeAgent Manual Test Checklist

- Load the add-in in Excel 2019 on Windows.
- Confirm the task pane renders the session sidebar and composer.
- Select `A1:D5` and verify the badge updates.
- Send a read-style prompt and verify no confirmation card appears.
- Send a write-style prompt and verify the confirmation card appears.
- Trigger `/upload_data 把选中数据上传到项目A` and verify the payload preview renders.
- Save an API key, reload the pane, and verify the key is restored.
- Create two sessions, switch between them, and verify histories remain isolated.
- Before release, replace the manifest `SourceLocation` with the production HTTPS task pane URL.
