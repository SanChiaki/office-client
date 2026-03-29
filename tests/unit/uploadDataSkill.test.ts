import { expect, test } from "vitest";
import { inferSkillRoute } from "../../src/skills/registry";
import { createUploadPreview } from "../../src/skills/uploadDataSkill";

test("matches upload intent from natural language", () => {
  expect(inferSkillRoute("把选中数据上传到项目A")).toEqual({
    skillName: "upload_data",
    project: "项目A",
  });
});

test("builds an upload preview from headers and rows", () => {
  expect(
    createUploadPreview("项目A", ["Name", "Owner"], [
      ["项目A", "张三"],
      ["项目B", "李四"],
      ["项目C", "王五"],
      ["项目D", "赵六"],
    ]),
  ).toEqual({
    project: "项目A",
    columns: ["Name", "Owner"],
    rowCount: 4,
    previewRows: [
      ["项目A", "张三"],
      ["项目B", "李四"],
      ["项目C", "王五"],
    ],
  });
});
