export interface ExcelAction {
  type: string;
  args: Record<string, unknown>;
}

export function classifyAction(action: ExcelAction) {
  return {
    requiresConfirmation: action.type.startsWith("excel.write") || action.type.includes("Sheet"),
  };
}

export function createExcelAdapter() {
  return {
    async readSelectionTable() {
      return {
        headers: ["Name", "Owner"],
        rows: [["\u9879\u76eeA", "\u5f20\u4e09"]],
      };
    },
    async run(action: ExcelAction) {
      return action.type;
    },
  };
}
