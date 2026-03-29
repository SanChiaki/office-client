import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { beforeEach, expect, test } from "vitest";
import App from "../../src/App";

beforeEach(() => {
  window.localStorage.clear();
  window.localStorage.setItem(
    "oa:settings",
    JSON.stringify({
      apiKey: "sk-demo",
      model: "gpt-4.1-mini",
    }),
  );
});

test("routes slash upload commands into the upload skill flow", async () => {
  const excelAdapter = {
    readSelectionTable: async () => ({
      headers: ["Name", "Owner"],
      rows: [["项目A", "张三"]],
    }),
    run: async (action: { type: string }) => action.type,
  };

  render(<App excelAdapterFactory={() => excelAdapter} />);

  const input = screen.getByRole("textbox", { name: "消息输入框" });
  const sendButton = document.querySelector(".composer-send");

  expect(sendButton).not.toBeNull();
  fireEvent.change(input, { target: { value: "/upload_data 把选中数据上传到项目A" } });
  fireEvent.click(sendButton!);

  await waitFor(() => {
    expect(screen.getByRole("button", { name: "确认" })).toBeInTheDocument();
  });

  expect(input).toHaveValue("");
  expect(
    within(screen.getByRole("region", { name: "消息线程" })).getByText("/upload_data 把选中数据上传到项目A"),
  ).toBeInTheDocument();
});
