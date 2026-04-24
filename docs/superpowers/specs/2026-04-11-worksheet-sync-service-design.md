Title: Worksheet Sync Service

## Context
- OfficeAgent.Core already provides `WorksheetChangeTracker` and `SyncOperationPreviewFactory`.
- Task 7 needs a coordinating `WorksheetSyncService` that speaks to the metadata store, connector, change tracker, and preview factory.
- Each ribbon action must read the latest metadata on every call; there is no caching at the service level.
- Missing bindings must surface errors so the UI can tell the user to bind a project; missing snapshots can be treated as empty.

## Requirements
1. `PrepareIncrementalUpload(string sheetName, IReadOnlyList<CellChange> currentCells)` must:
   - Load the latest worksheet snapshot from `IWorksheetMetadataStore`.
   - Determine dirty cells via `WorksheetChangeTracker`.
   - Skip cell changes without a `RowId`.
   - Only consider rows that already exist in the snapshot.
   - Return a preview via `SyncOperationPreviewFactory`. If no metadata or no dirty cells, return an empty preview (zero changes) and not throw.
2. `LoadSchemaForSheet(string sheetName)` should:
   - Load the sheet binding.
   - Throw if the binding is missing.
   - Ask `ISystemConnector.GetSchema(binding.ProjectId)` and return the result.
3. `ExecutePartialDownload(string sheetName, ResolvedSelection selection)` should:
   - Load binding (throw if absent).
   - Call `connector.Find(binding.ProjectId, selection.RowIds, selection.ApiFieldKeys)` and return the discovery result.
4. `ExecutePartialUpload(string sheetName, IReadOnlyList<CellChange> changes)` should:
   - Load binding (throw if absent).
   - Call `connector.BatchSave(binding.ProjectId, changes)`.

## Approaches
1. **Facade service (recommended)**: Build `WorksheetSyncService` with dependencies on the metadata store, connector, change tracker, and preview factory. Each method loads fresh metadata and orchestrates the collaborators as described above. This keeps orchestration logic centralized while keeping collaborators simple. **Trade-offs**: Some repeated metadata lookups, but the requirement explicitly forbids caching.
2. **Connector-aware service**: Let the connector carry more knowledge (e.g., fetching schema plus snapshot). This would blur responsibilities and duplicate tracker/preview logic in the connector. Not recommended.
3. **Command pattern**: Break each operation into command objects. Provides extensibility but adds ceremony for a single service; not worth extra complexity for Task 7.

## Design
- Implement `WorksheetSyncService` in `OfficeAgent.Core.Sync`.
- Constructor dependencies: `IWorksheetMetadataStore`, `ISystemConnector`, `WorksheetChangeTracker`, `SyncOperationPreviewFactory`.
- Each method must reload metadata via `IWorksheetMetadataStore` immediately before the operation. Throw `InvalidOperationException` (or `KeyNotFoundException`?) when a binding is missing.
- `PrepareIncrementalUpload` should treat a missing snapshot as an empty `WorksheetSnapshotCell[]`.
- Filtering logic:
  - Use `WorksheetChangeTracker.GetDirtyCells` with the latest snapshot.
  - After tracking, exclude entries where `RowId` is null/empty (per requirements).
  - Pass remaining changes to `SyncOperationPreviewFactory.CreateUploadPreview`. Use a consistent operation name like `"IncrementalUpload"`.
- Partial download/upload call signature should mirror what existing tests expect, returning connector results or void as appropriate.
- Provide tests covering:
  - Incremental upload orchestration returns preview for dirty existing rows without `RowId` and handles missing snapshot.
  - Schema load calls connector correctly and throws for missing binding.
  - Partial download and upload call connector with the binding project ID and propagate selection/changes.

## Testing Strategy
- Follow TDD: write the failing incremental upload test first.
- After service implementation, run `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`.
- Add targeted tests for each method to ensure metadata loading and connector usage behave as expected.

## Questions
- Confirmed earlier: metadata reload per call, binding missing ? throw, missing snapshot ? empty preview.
