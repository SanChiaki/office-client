import { afterEach, expect, test, vi } from "vitest";
import { requestCommandEnvelope } from "../../src/api/llmClient";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("posts the api key and payload to the agent endpoint", async () => {
  const json = vi.fn(async () => ({
    assistant_message: "ok",
    mode: "chat",
    requires_confirmation: false,
    actions: [],
  }));
  const fetchMock = vi.fn(async () => ({ ok: true, json }));
  vi.stubGlobal("fetch", fetchMock);

  const payload = { input: "hello" };

  const result = await requestCommandEnvelope("secret-key", payload);

  expect(result).toEqual({
    assistant_message: "ok",
    mode: "chat",
    requires_confirmation: false,
    actions: [],
  });
  expect(fetchMock).toHaveBeenCalledWith("https://api.example.com/agent", {
    method: "POST",
    headers: {
      Authorization: "Bearer secret-key",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
});

test("fails fast on non-2xx responses", async () => {
  const json = vi.fn();
  const fetchMock = vi.fn(async () => ({ ok: false, status: 500, json }));
  vi.stubGlobal("fetch", fetchMock);

  await expect(requestCommandEnvelope("secret-key", { input: "hello" })).rejects.toThrow("Request failed with status 500");
  expect(json).not.toHaveBeenCalled();
});

test("validates the response body before returning it", async () => {
  const fetchMock = vi.fn(async () => ({
    ok: true,
    json: vi.fn(async () => ({
      assistant_message: "ok",
      mode: "not-a-real-mode",
      requires_confirmation: false,
      actions: [],
    })),
  }));
  vi.stubGlobal("fetch", fetchMock);

  await expect(requestCommandEnvelope("secret-key", { input: "hello" })).rejects.toThrow();
});
