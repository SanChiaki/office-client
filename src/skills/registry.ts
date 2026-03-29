export interface SkillRoute {
  skillName: "upload_data";
  project: string;
}

function normalizeProject(project: string | undefined) {
  const normalized = (project ?? "").replace(/^把选中数据上传到/, "").trim();
  return normalized || "项目A";
}

export function inferSkillRoute(input: string): SkillRoute | null {
  const trimmed = input.trim();
  const slashMatch = trimmed.match(/^\/upload_data(?:\s+(.*))?$/);

  if (slashMatch) {
    return {
      skillName: "upload_data",
      project: normalizeProject(slashMatch[1]),
    };
  }

  if (trimmed.includes("上传") && trimmed.includes("项目")) {
    return {
      skillName: "upload_data",
      project: trimmed.match(/项目[^\s，。；,;]*/)?.[0] ?? "项目A",
    };
  }

  return null;
}
