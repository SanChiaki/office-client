export interface UploadPreview extends Record<string, unknown> {
  project: string;
  columns: string[];
  rowCount: number;
  previewRows: string[][];
}

export function buildUploadPreview(headers: string[], rows: string[][], project: string): UploadPreview {
  return {
    project,
    columns: headers,
    rowCount: rows.length,
    previewRows: rows.slice(0, 3),
  };
}
