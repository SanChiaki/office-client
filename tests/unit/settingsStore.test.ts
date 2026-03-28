import { beforeEach, expect, test } from "vitest";
import { createSettingsStore } from "../../src/state/settingsStore";

beforeEach(() => {
  window.localStorage.clear();
});

test("persists api key and model choice", () => {
  const store = createSettingsStore();
  store.save({ apiKey: "sk-demo", model: "gpt-4.1-mini" });
  expect(store.load()).toEqual({ apiKey: "sk-demo", model: "gpt-4.1-mini" });
});
