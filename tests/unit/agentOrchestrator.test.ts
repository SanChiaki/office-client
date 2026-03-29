import { expect, test } from "vitest";
import { decideRoute, shouldSendFullSelection } from "../../src/agent/agentOrchestrator";

test("routes /upload_data commands to the upload_data skill", () => {
  expect(decideRoute("/upload_data import this workbook")).toEqual({
    mode: "skill",
    skillName: "upload_data",
  });
});

test.each(["/upload_data", "/upload_data   next step"])("treats %s as the upload_data skill command", (input) => {
  expect(decideRoute(input)).toEqual({
    mode: "skill",
    skillName: "upload_data",
  });
});

test.each(["/upload_dataset", "/upload_datax", "prefix /upload_data"])("routes %s to chat", (input) => {
  expect(decideRoute(input)).toEqual({
    mode: "chat",
  });
});

test("routes natural-language upload requests to the upload_data skill", () => {
  expect(decideRoute("把选中数据上传到项目A")).toEqual({
    mode: "skill",
    skillName: "upload_data",
  });
});

test("sends full selection only for selections up to 25 cells", () => {
  expect(shouldSendFullSelection({ rowCount: 5, columnCount: 5 })).toBe(true);
  expect(shouldSendFullSelection({ rowCount: 6, columnCount: 5 })).toBe(false);
  expect(shouldSendFullSelection(null)).toBe(false);
});
