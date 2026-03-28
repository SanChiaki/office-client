import { fireEvent, render, screen, within } from "@testing-library/react";
import { beforeEach, expect, test } from "vitest";
import App from "../../src/App";

beforeEach(() => {
  window.localStorage.clear();
});

test("saves settings from the UI and rehydrates them on a fresh mount", () => {
  const { unmount } = render(<App />);

  fireEvent.click(screen.getByRole("button", { name: "设置" }));

  const dialog = screen.getByRole("dialog", { name: "Settings" });
  fireEvent.change(within(dialog).getByLabelText("API Key"), { target: { value: "sk-demo" } });
  fireEvent.change(within(dialog).getByLabelText("Model"), { target: { value: "gpt-4.1" } });
  fireEvent.click(within(dialog).getByRole("button", { name: "保存设置" }));

  unmount();
  render(<App />);

  fireEvent.click(screen.getByRole("button", { name: "设置" }));

  const rehydratedDialog = screen.getByRole("dialog", { name: "Settings" });
  expect(within(rehydratedDialog).getByLabelText("API Key")).toHaveValue("sk-demo");
  expect(within(rehydratedDialog).getByLabelText("Model")).toHaveValue("gpt-4.1");
});
