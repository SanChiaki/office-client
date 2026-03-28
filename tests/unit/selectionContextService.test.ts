import { expect, test } from "vitest";
import { normalizeSelection } from "../../src/excel/selectionContextService";

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
