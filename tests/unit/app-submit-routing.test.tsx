import { fireEvent, render, screen, within } from "@testing-library/react";
import { beforeEach, expect, test } from "vitest";
import App from "../../src/App";

beforeEach(() => {
  window.localStorage.clear();
});

test("keeps slash-command drafts intact until skill execution exists", () => {
  render(<App />);

  const input = screen.getByRole("textbox", { name: "消息输入框" });
  const sendButton = document.querySelector(".composer-send");

  expect(sendButton).not.toBeNull();
  fireEvent.change(input, { target: { value: "/upload_data import this workbook" } });
  fireEvent.click(sendButton!);

  expect(input).toHaveValue("/upload_data import this workbook");
  expect(within(screen.getByRole("region", { name: "消息线程" })).queryByText("/upload_data import this workbook")).not.toBeInTheDocument();
});
