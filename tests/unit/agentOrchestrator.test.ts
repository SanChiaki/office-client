import { expect, test } from "vitest";
import { decideRoute } from "../../src/agent/agentOrchestrator";

test("routes /upload_data commands to the upload_data skill", () => {
  expect(decideRoute("/upload_data import this workbook")).toEqual({
    mode: "skill",
    skillName: "upload_data",
  });
});
