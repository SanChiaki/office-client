import { afterEach, expect, test, vi } from "vitest";

import { uploadData } from "../../src/api/businessApiClient";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("posts upload data to the configured base url", async () => {
  const fetchMock = vi.fn(async () => ({
    ok: true,
    json: vi.fn(async () => ({ saved: 1 })),
  }));

  vi.stubGlobal("fetch", fetchMock);

  await uploadData("sk-demo", "https://internal.example/api/", {
    project: "Project A",
  });

  expect(fetchMock).toHaveBeenCalledWith("https://internal.example/api/upload_data_api", {
    method: "POST",
    headers: {
      Authorization: "Bearer sk-demo",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      project: "Project A",
    }),
  });
});
