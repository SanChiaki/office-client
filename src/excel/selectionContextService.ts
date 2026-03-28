import type { SelectionContext } from "../types";

export interface RawSelection {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}

type SelectionRange = {
  address?: string;
  rowCount?: number;
  columnCount?: number;
  load?: (properties: string[] | string) => void;
  worksheet?: {
    name?: string;
    load?: (properties: string[] | string) => void;
  };
};

type SelectionContextObject = {
  workbook: {
    getSelectedRange: () => SelectionRange;
  };
  sync: () => Promise<void>;
};

type ExcelRuntime = {
  run?: <T>(callback: (context: SelectionContextObject) => Promise<T> | T) => Promise<T>;
};

type OfficeDocument = {
  addHandlerAsync?: (
    eventType: string,
    handler: () => Promise<void> | void,
    callback?: (result: { status: string }) => void,
  ) => void;
  removeHandlerAsync?: (
    eventType: string,
    handler: { handler?: () => Promise<void> | void } | (() => Promise<void> | void),
    callback?: (result: { status: string }) => void,
  ) => void;
};

type OfficeRuntime = {
  onReady?: () => Promise<unknown> | void;
  context?: {
    document?: OfficeDocument;
  };
};

export function normalizeSelection(raw: RawSelection): SelectionContext {
  return {
    ...raw,
    hasHeaders: false,
  };
}

async function readCurrentSelection(): Promise<SelectionContext | null> {
  const runtime = window as unknown as { Excel?: ExcelRuntime };
  if (!runtime.Excel?.run) {
    return null;
  }

  return runtime.Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load?.(["address", "rowCount", "columnCount"]);
    range.worksheet?.load?.("name");
    await context.sync();

    const sheetName = range.worksheet?.name;
    const address = range.address;
    const rowCount = range.rowCount;
    const columnCount = range.columnCount;

    if (
      typeof sheetName !== "string" ||
      typeof address !== "string" ||
      typeof rowCount !== "number" ||
      typeof columnCount !== "number"
    ) {
      return null;
    }

    return normalizeSelection({
      sheetName,
      address,
      rowCount,
      columnCount,
    });
  });
}

export function subscribeToSelectionChanges(onChange: (selection: SelectionContext) => void) {
  const office = (window as unknown as { Office?: OfficeRuntime }).Office;
  let disposed = false;
  let handler: (() => Promise<void> | void) | null = null;
  let removeHandlerAsync: OfficeDocument["removeHandlerAsync"] | null = null;

  async function emitCurrentSelection() {
    try {
      const selection = await readCurrentSelection();
      if (!disposed && selection) {
        onChange(selection);
      }
    } catch {
      // Selection refresh is best-effort. Ignore runtime failures for now.
    }
  }

  function register() {
    if (disposed) {
      return;
    }

    const document = (window as unknown as { Office?: OfficeRuntime }).Office?.context?.document;
    if (!document?.addHandlerAsync || !document?.removeHandlerAsync) {
      return;
    }

    removeHandlerAsync = document.removeHandlerAsync;
    handler = async () => {
      await emitCurrentSelection();
    };

    document.addHandlerAsync("documentSelectionChanged", handler, () => {
      void emitCurrentSelection();
    });
  }

  const ready = office?.onReady?.();
  if (ready && typeof (ready as Promise<unknown>).then === "function") {
    void (ready as Promise<unknown>).then(register).catch(() => {});
  } else {
    register();
  }

  return () => {
    disposed = true;
    if (handler && removeHandlerAsync) {
      removeHandlerAsync("documentSelectionChanged", { handler }, () => {});
    }
  };
}
