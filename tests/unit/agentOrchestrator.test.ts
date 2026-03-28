import { expect, test } from "vitest";
import { decideRoute } from "../../src/agent/agentOrchestrator";

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
