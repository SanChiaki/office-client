import { buildUploadPreview } from "./uploadPayloadBuilder";

export function createUploadPreview(project: string, headers: string[], rows: string[][]) {
  return buildUploadPreview(headers, rows, project);
}
