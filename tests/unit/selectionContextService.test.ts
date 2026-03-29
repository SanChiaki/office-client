import { afterEach, expect, test, vi } from "vitest";
import {
  normalizeSelection,
  shouldUseSummaryMode,
  subscribeToSelectionChanges,
} from "../../src/excel/selectionContextService";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("normalizes raw excel selection metadata", () => {
  expect(
    normalizeSelection({
      sheetName: "Sheet1",
      address: "A1:D4",
      rowCount: 4,
      columnCount: 4,
    }),
  ).toEqual({
    sheetName: "Sheet1",
    address: "A1:D4",
    rowCount: 4,
    columnCount: 4,
    hasHeaders: false,
  });
});

test("uses summary mode for selections larger than 25 cells", () => {
  expect(shouldUseSummaryMode({ rowCount: 6, columnCount: 5 })).toBe(true);
  expect(shouldUseSummaryMode({ rowCount: 5, columnCount: 5 })).toBe(false);
});

test("hydrates the current selection immediately after registration", async () => {
  let registeredHandler: ((eventArgs: unknown) => Promise<void> | void) | null = null;
  const addHandlerAsync = vi.fn((eventType: string, handler: typeof registeredHandler, callback?: (result: { status: string }) => void) => {
    registeredHandler = handler;
    callback?.({ status: "succeeded" });
  });
  const removeHandlerAsync = vi.fn();
  const load = vi.fn();
  const worksheetLoad = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "A1:D4",
    rowCount: 4,
    columnCount: 4,
    load,
    worksheet: {
      load: worksheetLoad,
      name: "Sheet1",
    },
  }));
  const run = vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) => {
    return callback({
      workbook: { getSelectedRange },
      sync,
    });
  });

  vi.stubGlobal("Office", {
    onReady: () => Promise.resolve(),
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  });
  vi.stubGlobal("Excel", { run });

  const onChange = vi.fn();
  subscribeToSelectionChanges(onChange);

  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(onChange).toHaveBeenCalledWith({
    sheetName: "Sheet1",
    address: "A1:D4",
    rowCount: 4,
    columnCount: 4,
    hasHeaders: false,
  });

  expect(addHandlerAsync).toHaveBeenCalledTimes(1);
  expect(getSelectedRange).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledWith(["address", "rowCount", "columnCount"]);
  expect(worksheetLoad).toHaveBeenCalledWith("name");
  expect(sync).toHaveBeenCalledTimes(1);
  expect(removeHandlerAsync).not.toHaveBeenCalled();
  expect(registeredHandler).toEqual(expect.any(Function));
});

test("waits for Office.onReady before registering and still updates", async () => {
  type SelectionHandler = (eventArgs: unknown) => Promise<void> | void;

  let resolveReady!: () => void;
  const ready = new Promise<void>((resolve) => {
    resolveReady = resolve;
  });
  let registeredHandler: SelectionHandler | null = null;
  const addHandlerAsync = vi.fn((eventType: string, handler: SelectionHandler, callback?: (result: { status: string }) => void) => {
    registeredHandler = handler;
    callback?.({ status: "succeeded" });
  });
  const removeHandlerAsync = vi.fn();
  const load = vi.fn();
  const worksheetLoad = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    load,
    worksheet: {
      load: worksheetLoad,
      name: "Budget",
    },
  }));
  const run = vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) => {
    return callback({
      workbook: { getSelectedRange },
      sync,
    });
  });

  const office = {
    onReady: () => ready,
    context: undefined as
      | {
          document: {
            addHandlerAsync: typeof addHandlerAsync;
            removeHandlerAsync: typeof removeHandlerAsync;
          };
        }
      | undefined,
  };

  vi.stubGlobal("Office", office);
  vi.stubGlobal("Excel", { run });

  const onChange = vi.fn();
  const cleanup = subscribeToSelectionChanges(onChange);

  expect(addHandlerAsync).not.toHaveBeenCalled();

  office.context = {
    document: {
      addHandlerAsync,
      removeHandlerAsync,
    },
  };
  resolveReady();

  await new Promise((resolve) => setTimeout(resolve, 0));

  expect(addHandlerAsync).toHaveBeenCalledTimes(1);
  expect(onChange).toHaveBeenCalledWith({
    sheetName: "Budget",
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    hasHeaders: false,
  });

  const selectionHandler = registeredHandler;
  expect(selectionHandler).toEqual(expect.any(Function));
  if (!selectionHandler) {
    throw new Error("Selection handler was not registered");
  }

  await (selectionHandler as SelectionHandler)({ document: {} });

  expect(onChange).toHaveBeenCalledWith({
    sheetName: "Budget",
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    hasHeaders: false,
  });
  expect(removeHandlerAsync).not.toHaveBeenCalled();
  cleanup();
});

test("registers a selection handler, emits normalized metadata, and cleans up", async () => {
  type SelectionHandler = (eventArgs: unknown) => Promise<void> | void;

  let registeredHandler: SelectionHandler | null = null;
  const removeHandlerAsync = vi.fn((eventType: string, payload: { handler: SelectionHandler | null }, callback?: (result: { status: string }) => void) => {
    callback?.({ status: "succeeded" });
  });
  const addHandlerAsync = vi.fn((eventType: string, handler: SelectionHandler, callback?: (result: { status: string }) => void) => {
    registeredHandler = handler;
    callback?.({ status: "succeeded" });
  });
  const load = vi.fn();
  const worksheetLoad = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    load,
    worksheet: {
      load: worksheetLoad,
      name: "Budget",
    },
  }));
  const run = vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) => {
    return callback({
      workbook: { getSelectedRange },
      sync,
    });
  });

  vi.stubGlobal("Office", {
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  });
  vi.stubGlobal("Excel", { run });

  const onChange = vi.fn();
  const cleanup = subscribeToSelectionChanges(onChange);

  expect(addHandlerAsync).toHaveBeenCalledWith(
    "documentSelectionChanged",
    expect.any(Function),
    expect.any(Function),
  );

  const selectionHandler = registeredHandler;
  expect(selectionHandler).toEqual(expect.any(Function));
  if (!selectionHandler) {
    throw new Error("Selection handler was not registered");
  }

  await (selectionHandler as SelectionHandler)({ document: {} });

  expect(run).toHaveBeenCalledTimes(2);
  expect(getSelectedRange).toHaveBeenCalledTimes(2);
  expect(load).toHaveBeenCalledWith(["address", "rowCount", "columnCount"]);
  expect(worksheetLoad).toHaveBeenCalledWith("name");
  expect(sync).toHaveBeenCalledTimes(2);
  expect(onChange).toHaveBeenNthCalledWith(1, {
    sheetName: "Budget",
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    hasHeaders: false,
  });
  expect(onChange).toHaveBeenNthCalledWith(2, {
    sheetName: "Budget",
    address: "B2:C3",
    rowCount: 2,
    columnCount: 2,
    hasHeaders: false,
  });

  cleanup();

  expect(removeHandlerAsync).toHaveBeenCalledWith(
    "documentSelectionChanged",
    { handler: selectionHandler },
    expect.any(Function),
  );
});
