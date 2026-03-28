import { expect, test } from "vitest";
import { classifyAction } from "../../src/excel/excelAdapter";

test("requires confirmation for excel.writeRange actions", () => {
  expect(classifyAction({ type: "excel.writeRange", args: {} })).toEqual({
    requiresConfirmation: true,
  });
});
