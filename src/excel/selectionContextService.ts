import type { SelectionContext } from "../types";

export interface RawSelection {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}

export function normalizeSelection(raw: RawSelection): SelectionContext {
  return {
    ...raw,
    hasHeaders: false,
  };
}

export function subscribeToSelectionChanges(onChange: (selection: SelectionContext) => void) {
  const office = (window as unknown as { Office?: any }).Office;

  if (!office?.context?.document?.addHandlerAsync) {
    return () => {};
  }

  office.context.document.addHandlerAsync("documentSelectionChanged", async () => {
    onChange(
      normalizeSelection({
        sheetName: "Sheet1",
        address: "A1",
        rowCount: 1,
        columnCount: 1,
      }),
    );
  });

  return () => {};
}
